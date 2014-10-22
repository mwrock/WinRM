# Copyright 2010 Dan Wanek <dan.wanek@gmail.com>
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#     http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

require 'nori'
require 'rexml/document'
require 'zip'

module WinRM
  # This is the main class that does the SOAP request/response logic. There are a few helper classes, but pretty
  #   much everything comes through here first.
  class WinRMWebService

    DEFAULT_TIMEOUT = 'PT60S'
    DEFAULT_MAX_ENV_SIZE = 153600
    DEFAULT_LOCALE = 'en-US'

    attr_reader :endpoint, :timeout

    # @param [String,URI] endpoint the WinRM webservice endpoint
    # @param [Symbol] transport either :kerberos(default)/:ssl/:plaintext
    # @param [Hash] opts Misc opts for the various transports.
    #   @see WinRM::HTTP::HttpTransport
    #   @see WinRM::HTTP::HttpGSSAPI
    #   @see WinRM::HTTP::HttpSSL
    def initialize(endpoint, transport = :kerberos, opts = {})
      @endpoint = endpoint
      @timeout = DEFAULT_TIMEOUT
      @max_env_sz = DEFAULT_MAX_ENV_SIZE
      @locale = DEFAULT_LOCALE
      @logger = Logging.logger[self]
      case transport
      when :kerberos
        require 'gssapi'
        @xfer = HTTP::HttpGSSAPI.new(endpoint, opts[:realm], opts[:service], opts[:keytab], opts)
      when :plaintext
        @xfer = HTTP::HttpPlaintext.new(endpoint, opts[:user], opts[:pass], opts)
      when :ssl
        @xfer = HTTP::HttpSSL.new(endpoint, opts[:user], opts[:pass], opts[:ca_trust_path], opts)
      else
        raise "Invalid transport '#{transport}' specified, expected: kerberos, plaintext, ssl."
      end
    end

    # Operation timeout
    # @see http://msdn.microsoft.com/en-us/library/ee916629(v=PROT.13).aspx
    # @param [Fixnum] sec the number of seconds to set the timeout to. It will be converted to an ISO8601 format.
    def set_timeout(sec)
      @timeout = Iso8601Duration.sec_to_dur(sec)
    end
    alias :op_timeout :set_timeout

    # Max envelope size
    # @see http://msdn.microsoft.com/en-us/library/ee916127(v=PROT.13).aspx
    # @param [Fixnum] byte_sz the max size in bytes to allow for the response
    def max_env_size(byte_sz)
      @max_env_sz = byte_sz
    end

    # Set the locale
    # @see http://msdn.microsoft.com/en-us/library/gg567404(v=PROT.13).aspx
    # @param [String] locale the locale to set for future messages
    def locale(locale)
      @locale = locale
    end

    # Create a Shell on the destination host
    # @param [Hash<optional>] shell_opts additional shell options you can pass
    # @option shell_opts [String] :i_stream Which input stream to open.  Leave this alone unless you know what you're doing (default: stdin)
    # @option shell_opts [String] :o_stream Which output stream to open.  Leave this alone unless you know what you're doing (default: stdout stderr)
    # @option shell_opts [String] :working_directory the directory to create the shell in
    # @option shell_opts [Hash] :env_vars environment variables to set for the shell. For instance;
    #   :env_vars => {:myvar1 => 'val1', :myvar2 => 'var2'}
    # @return [String] The ShellId from the SOAP response.  This is our open shell instance on the remote machine.
    def open_shell(shell_opts = {})
      i_stream = shell_opts.has_key?(:i_stream) ? shell_opts[:i_stream] : 'stdin'
      o_stream = shell_opts.has_key?(:o_stream) ? shell_opts[:o_stream] : 'stdout stderr'
      codepage = shell_opts.has_key?(:codepage) ? shell_opts[:codepage] : 437
      noprofile = shell_opts.has_key?(:noprofile) ? shell_opts[:noprofile] : 'FALSE'
      h_opts = { "#{NS_WSMAN_DMTF}:OptionSet" => { "#{NS_WSMAN_DMTF}:Option" => [noprofile, codepage],
        :attributes! => {"#{NS_WSMAN_DMTF}:Option" => {'Name' => ['WINRS_NOPROFILE','WINRS_CODEPAGE']}}}}
      shell_body = {
        "#{NS_WIN_SHELL}:InputStreams" => i_stream,
        "#{NS_WIN_SHELL}:OutputStreams" => o_stream
      }
      shell_body["#{NS_WIN_SHELL}:WorkingDirectory"] = shell_opts[:working_directory] if shell_opts.has_key?(:working_directory)
      shell_body["#{NS_WIN_SHELL}:IdleTimeOut"] = shell_opts[:idle_timeout] if(shell_opts.has_key?(:idle_timeout) && shell_opts[:idle_timeout].is_a?(String))
      if(shell_opts.has_key?(:env_vars) && shell_opts[:env_vars].is_a?(Hash))
        keys = shell_opts[:env_vars].keys
        vals = shell_opts[:env_vars].values
        shell_body["#{NS_WIN_SHELL}:Environment"] = {
          "#{NS_WIN_SHELL}:Variable" => vals,
          :attributes! => {"#{NS_WIN_SHELL}:Variable" => {'Name' => keys}}
        }
      end
      builder = Builder::XmlMarkup.new
      builder.instruct!(:xml, :encoding => 'UTF-8')
      builder.tag! :env, :Envelope, namespaces do |env|
        env.tag!(:env, :Header) { |h| h << Gyoku.xml(merge_headers(header,resource_uri_cmd,action_create,h_opts)) }
        env.tag! :env, :Body do |body|
          body.tag!("#{NS_WIN_SHELL}:Shell") { |s| s << Gyoku.xml(shell_body)}
        end
      end

      resp_doc = send_message(builder.target!)
      shell_id = REXML::XPath.first(resp_doc, "//*[@Name='ShellId']").text
      shell_id
    end

    # Run a command on a machine with an open shell
    # @param [String] shell_id The shell id on the remote machine.  See #open_shell
    # @param [String] command The command to run on the remote machine
    # @param [Array<String>] arguments An array of arguments for this command
    # @return [String] The CommandId from the SOAP response.  This is the ID we need to query in order to get output.
    def run_command(shell_id, command, arguments = [], cmd_opts = {})
      consolemode = cmd_opts.has_key?(:console_mode_stdin) ? cmd_opts[:console_mode_stdin] : 'TRUE'
      skipcmd     = cmd_opts.has_key?(:skip_cmd_shell) ? cmd_opts[:skip_cmd_shell] : 'FALSE'

      h_opts = { "#{NS_WSMAN_DMTF}:OptionSet" => {
        "#{NS_WSMAN_DMTF}:Option" => [consolemode, skipcmd],
        :attributes! => {"#{NS_WSMAN_DMTF}:Option" => {'Name' => ['WINRS_CONSOLEMODE_STDIN','WINRS_SKIP_CMD_SHELL']}}}
      }
      body = { "#{NS_WIN_SHELL}:Command" => "\"#{command}\"", "#{NS_WIN_SHELL}:Arguments" => arguments}

      builder = Builder::XmlMarkup.new
      builder.instruct!(:xml, :encoding => 'UTF-8')
      builder.tag! :env, :Envelope, namespaces do |env|
        env.tag!(:env, :Header) { |h| h << Gyoku.xml(merge_headers(header,resource_uri_cmd,action_command,h_opts,selector_shell_id(shell_id))) }
        env.tag!(:env, :Body) do |env_body|
          env_body.tag!("#{NS_WIN_SHELL}:CommandLine") { |cl| cl << Gyoku.xml(body) }
        end
      end

      # Grab the command element and unescape any single quotes - issue 69
      xml = builder.target!
      escaped_cmd = /<#{NS_WIN_SHELL}:Command>(.+)<\/#{NS_WIN_SHELL}:Command>/.match(xml)[1]
      xml.sub!(escaped_cmd, escaped_cmd.gsub(/&#39;/, "'"))

      resp_doc = send_message(xml)
      command_id = REXML::XPath.first(resp_doc, "//#{NS_WIN_SHELL}:CommandId").text
      command_id
    end

    def upload_path(local, remote)
      @logger.debug("Upload: #{local} -> #{remote}")
      local = Array.new(1) { local } if local.is_a? String
      shell_id = open_shell
      local.each do |path|
        if File.directory?(path)
          upload_directory(shell_id, path, remote)
        else
          upload_file(path, File.join(remote, File.basename(path)), shell_id)
        end
      end
    ensure
      close_shell(shell_id)
    end

    def write_stdin(shell_id, command_id, stdin)
      # Signal the Command references to terminate (close stdout/stderr)
      body = {
        "#{NS_WIN_SHELL}:Send" => {
          "#{NS_WIN_SHELL}:Stream" => {
            "@Name" => 'stdin',
            "@CommandId" => command_id,
            :content! => Base64.encode64(stdin)
          }
        }
      }
      builder = Builder::XmlMarkup.new
      builder.instruct!(:xml, :encoding => 'UTF-8')
      builder.tag! :env, :Envelope, namespaces do |env|
        env.tag!(:env, :Header) { |h| h << Gyoku.xml(merge_headers(header,resource_uri_cmd,action_send,selector_shell_id(shell_id))) }
        env.tag!(:env, :Body) do |env_body|
          env_body << Gyoku.xml(body)
        end
      end
      resp = send_message(builder.target!)
      true
    end

    # Get the Output of the given shell and command
    # @param [String] shell_id The shell id on the remote machine.  See #open_shell
    # @param [String] command_id The command id on the remote machine.  See #run_command
    # @return [Hash] Returns a Hash with a key :exitcode and :data.  Data is an Array of Hashes where the cooresponding key
    #   is either :stdout or :stderr.  The reason it is in an Array so so we can get the output in the order it ocurrs on
    #   the console.
    def get_command_output(shell_id, command_id, &block)
      body = { "#{NS_WIN_SHELL}:DesiredStream" => 'stdout stderr',
        :attributes! => {"#{NS_WIN_SHELL}:DesiredStream" => {'CommandId' => command_id}}}

      builder = Builder::XmlMarkup.new
      builder.instruct!(:xml, :encoding => 'UTF-8')
      builder.tag! :env, :Envelope, namespaces do |env|
        env.tag!(:env, :Header) { |h| h << Gyoku.xml(merge_headers(header,resource_uri_cmd,action_receive,selector_shell_id(shell_id))) }
        env.tag!(:env, :Body) do |env_body|
          env_body.tag!("#{NS_WIN_SHELL}:Receive") { |cl| cl << Gyoku.xml(body) }
        end
      end

      resp_doc = send_message(builder.target!)
      output = {:data => []}
      REXML::XPath.match(resp_doc, "//#{NS_WIN_SHELL}:Stream").each do |n|
        next if n.text.nil? || n.text.empty?
        stream = { n.attributes['Name'].to_sym => Base64.decode64(n.text) }
        output[:data] << stream
        yield stream[:stdout], stream[:stderr] if block_given?
      end

      # We may need to get additional output if the stream has not finished.
      # The CommandState will change from Running to Done like so:
      # @example
      #   from...
      #   <rsp:CommandState CommandId="..." State="http://schemas.microsoft.com/wbem/wsman/1/windows/shell/CommandState/Running"/>
      #   to...
      #   <rsp:CommandState CommandId="..." State="http://schemas.microsoft.com/wbem/wsman/1/windows/shell/CommandState/Done">
      #     <rsp:ExitCode>0</rsp:ExitCode>
      #   </rsp:CommandState>
      #if ((resp/"//*[@State='http://schemas.microsoft.com/wbem/wsman/1/windows/shell/CommandState/Done']").empty?)
      done_elems = REXML::XPath.match(resp_doc, "//*[@State='http://schemas.microsoft.com/wbem/wsman/1/windows/shell/CommandState/Done']")
      if done_elems.empty?
        output.merge!(get_command_output(shell_id, command_id, &block)) do |key, old_data, new_data|
          old_data += new_data
        end
      else
        output[:exitcode] = REXML::XPath.first(resp_doc, "//#{NS_WIN_SHELL}:ExitCode").text.to_i
      end
      output
    end

    # Clean-up after a command.
    # @see #run_command
    # @param [String] shell_id The shell id on the remote machine.  See #open_shell
    # @param [String] command_id The command id on the remote machine.  See #run_command
    # @return [true] This should have more error checking but it just returns true for now.
    def cleanup_command(shell_id, command_id)
      # Signal the Command references to terminate (close stdout/stderr)
      body = { "#{NS_WIN_SHELL}:Code" => 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell/signal/terminate' }
      builder = Builder::XmlMarkup.new
      builder.instruct!(:xml, :encoding => 'UTF-8')
      builder.tag! :env, :Envelope, namespaces do |env|
        env.tag!(:env, :Header) { |h| h << Gyoku.xml(merge_headers(header,resource_uri_cmd,action_signal,selector_shell_id(shell_id))) }
        env.tag!(:env, :Body) do |env_body|
          env_body.tag!("#{NS_WIN_SHELL}:Signal", {'CommandId' => command_id}) { |cl| cl << Gyoku.xml(body) }
        end
      end
      resp = send_message(builder.target!)
      true
    end

    # Close the shell
    # @param [String] shell_id The shell id on the remote machine.  See #open_shell
    # @return [true] This should have more error checking but it just returns true for now.
    def close_shell(shell_id)
      builder = Builder::XmlMarkup.new
      builder.instruct!(:xml, :encoding => 'UTF-8')

      builder.tag!('env:Envelope', namespaces) do |env|
        env.tag!('env:Header') { |h| h << Gyoku.xml(merge_headers(header,resource_uri_cmd,action_delete,selector_shell_id(shell_id))) }
        env.tag!('env:Body')
      end

      resp = send_message(builder.target!)
      true
    end

    # Run a CMD command
    # @param [String] command The command to run on the remote system
    # @param [Array <String>] arguments arguments to the command
    # @param [String] an existing and open shell id to reuse
    # @return [Hash] :stdout and :stderr
    def run_cmd(command, arguments = [], shell_id = nil, &block)
      shell_given = !shell_id.nil?
      shell_id ||= open_shell
      command_id =  run_command(shell_id, command, arguments)
      command_output = get_command_output(shell_id, command_id, &block)
      cleanup_command(shell_id, command_id)
      close_shell(shell_id) unless shell_given
      command_output
    end
    alias :cmd :run_cmd


    # Run a Powershell script that resides on the local box.
    # @param [IO,String] script_file an IO reference for reading the Powershell script or the actual file contents
    # @param [String] an existing and open shell id to reuse
    # @return [Hash] :stdout and :stderr
    def run_powershell_script(script_file, shell_id = nil, &block)
      shell_given = !shell_id.nil?
      # if an IO object is passed read it..otherwise assume the contents of the file were passed
      script = script_file.kind_of?(IO) ? script_file.read : script_file
      script = script.encode('UTF-16LE', 'UTF-8')
      script = Base64.strict_encode64(script)
      shell_id ||= open_shell
      command_id = run_command(shell_id, "powershell -encodedCommand #{script}")
      command_output = get_command_output(shell_id, command_id, &block)
      cleanup_command(shell_id, command_id)
      close_shell(shell_id) unless shell_given
      command_output
    end
    alias :powershell :run_powershell_script


    # Run a WQL Query
    # @see http://msdn.microsoft.com/en-us/library/aa394606(VS.85).aspx
    # @param [String] wql The WQL query
    # @return [Hash] Returns a Hash that contain the key/value pairs returned from the query.
    def run_wql(wql)

      body = { "#{NS_WSMAN_DMTF}:OptimizeEnumeration" => nil,
        "#{NS_WSMAN_DMTF}:MaxElements" => '32000',
        "#{NS_WSMAN_DMTF}:Filter" => wql,
        :attributes! => { "#{NS_WSMAN_DMTF}:Filter" => {'Dialect' => 'http://schemas.microsoft.com/wbem/wsman/1/WQL'}}
      }

      builder = Builder::XmlMarkup.new
      builder.instruct!(:xml, :encoding => 'UTF-8')
      builder.tag! :env, :Envelope, namespaces do |env|
        env.tag!(:env, :Header) { |h| h << Gyoku.xml(merge_headers(header,resource_uri_wmi,action_enumerate)) }
        env.tag!(:env, :Body) do |env_body|
          env_body.tag!("#{NS_ENUM}:Enumerate") { |en| en << Gyoku.xml(body) }
        end
      end

      resp = send_message(builder.target!)
      parser = Nori.new(:parser => :rexml, :advanced_typecasting => false, :convert_tags_to => lambda { |tag| tag.snakecase.to_sym }, :strip_namespaces => true)
      hresp = parser.parse(resp.to_s)[:envelope][:body]
      
      # Normalize items so the type always has an array even if it's just a single item.
      items = {}
      if hresp[:enumerate_response][:items]
        hresp[:enumerate_response][:items].each_pair do |k,v|
          if v.is_a?(Array)
            items[k] = v
          else
            items[k] = [v]
          end
        end
      end
      items
    end
    alias :wql :run_wql

    def toggle_nori_type_casting(to)
      @logger.warn('toggle_nori_type_casting is deprecated and has no effect, ' +
        'please remove calls to it')
    end

    private

    def namespaces
      {
        'xmlns:xsd' => 'http://www.w3.org/2001/XMLSchema',
        'xmlns:xsi' => 'http://www.w3.org/2001/XMLSchema-instance',
        'xmlns:env' => 'http://www.w3.org/2003/05/soap-envelope',
        'xmlns:a' => 'http://schemas.xmlsoap.org/ws/2004/08/addressing',
        'xmlns:b' => 'http://schemas.dmtf.org/wbem/wsman/1/cimbinding.xsd',
        'xmlns:n' => 'http://schemas.xmlsoap.org/ws/2004/09/enumeration',
        'xmlns:x' => 'http://schemas.xmlsoap.org/ws/2004/09/transfer',
        'xmlns:w' => 'http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd',
        'xmlns:p' => 'http://schemas.microsoft.com/wbem/wsman/1/wsman.xsd',
        'xmlns:rsp' => 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell',
        'xmlns:cfg' => 'http://schemas.microsoft.com/wbem/wsman/1/config',
      }
    end

    def header
      { "#{NS_ADDRESSING}:To" => "#{@xfer.endpoint.to_s}",
        "#{NS_ADDRESSING}:ReplyTo" => {
        "#{NS_ADDRESSING}:Address" => 'http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous',
          :attributes! => {"#{NS_ADDRESSING}:Address" => {'mustUnderstand' => true}}},
        "#{NS_WSMAN_DMTF}:MaxEnvelopeSize" => @max_env_sz,
        "#{NS_ADDRESSING}:MessageID" => "uuid:#{UUIDTools::UUID.random_create.to_s.upcase}",
        "#{NS_WSMAN_DMTF}:Locale/" => '',
        "#{NS_WSMAN_MSFT}:DataLocale/" => '',
        "#{NS_WSMAN_DMTF}:OperationTimeout" => @timeout,
        :attributes! => {
          "#{NS_WSMAN_DMTF}:MaxEnvelopeSize" => {'mustUnderstand' => true},
          "#{NS_WSMAN_DMTF}:Locale/" => {'xml:lang' => @locale, 'mustUnderstand' => false},
          "#{NS_WSMAN_MSFT}:DataLocale/" => {'xml:lang' => @locale, 'mustUnderstand' => false}
        }}
    end

    # merge the various header hashes and make sure we carry all of the attributes
    #   through instead of overwriting them.
    def merge_headers(*headers)
      hdr = {}
      headers.each do |h|
        hdr.merge!(h) do |k,v1,v2|
          v1.merge!(v2) if k == :attributes!
        end
      end
      hdr
    end

    def send_message(message)
      @xfer.send_request(message)
    end
    

    # Helper methods for SOAP Headers

    def resource_uri_cmd
      {"#{NS_WSMAN_DMTF}:ResourceURI" => 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd',
        :attributes! => {"#{NS_WSMAN_DMTF}:ResourceURI" => {'mustUnderstand' => true}}}
    end

    def resource_uri_wmi(namespace = 'root/cimv2/*')
      {"#{NS_WSMAN_DMTF}:ResourceURI" => "http://schemas.microsoft.com/wbem/wsman/1/wmi/#{namespace}",
        :attributes! => {"#{NS_WSMAN_DMTF}:ResourceURI" => {'mustUnderstand' => true}}}
    end

    def action_create
      {"#{NS_ADDRESSING}:Action" => 'http://schemas.xmlsoap.org/ws/2004/09/transfer/Create',
        :attributes! => {"#{NS_ADDRESSING}:Action" => {'mustUnderstand' => true}}}
    end

    def action_delete
      {"#{NS_ADDRESSING}:Action" => 'http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete',
        :attributes! => {"#{NS_ADDRESSING}:Action" => {'mustUnderstand' => true}}}
    end

    def action_command
      {"#{NS_ADDRESSING}:Action" => 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command',
        :attributes! => {"#{NS_ADDRESSING}:Action" => {'mustUnderstand' => true}}}
    end

    def action_receive
      {"#{NS_ADDRESSING}:Action" => 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive',
        :attributes! => {"#{NS_ADDRESSING}:Action" => {'mustUnderstand' => true}}}
    end

    def action_signal
      {"#{NS_ADDRESSING}:Action" => 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Signal',
        :attributes! => {"#{NS_ADDRESSING}:Action" => {'mustUnderstand' => true}}}
    end

    def action_send
      {"#{NS_ADDRESSING}:Action" => 'http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Send',
        :attributes! => {"#{NS_ADDRESSING}:Action" => {'mustUnderstand' => true}}}
    end

    def action_enumerate
      {"#{NS_ADDRESSING}:Action" => 'http://schemas.xmlsoap.org/ws/2004/09/enumeration/Enumerate',
        :attributes! => {"#{NS_ADDRESSING}:Action" => {'mustUnderstand' => true}}}
    end

    def selector_shell_id(shell_id)
      {"#{NS_WSMAN_DMTF}:SelectorSet" =>
        {"#{NS_WSMAN_DMTF}:Selector" => shell_id, :attributes! => {"#{NS_WSMAN_DMTF}:Selector" => {'Name' => 'ShellId'}}}
      }
    end

    def upload_file(local, remote, shell_id)
      @logger.debug("Upload: #{local} -> #{remote}")
      remote = full_remote_path(remote, shell_id)
      if should_upload_file?(shell_id, local, remote)
        upload_to_remote(shell_id, local, remote)
        decode_remote_file(shell_id, local, remote)
      end
    end

    def full_remote_path(path, shell_id)
      command = <<-EOH
        $dest_file_path = [System.IO.Path]::GetFullPath('#{path}')

        if (!(Test-Path $dest_file_path)) {
          $dest_dir = ([System.IO.Path]::GetDirectoryName($dest_file_path))
          New-Item -ItemType directory -Force -Path $dest_dir | Out-Null
        }

        $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath("#{path}")
      EOH

      powershell(command, shell_id)[:data][0][:stdout].chomp
    end

    def should_upload_file?(shell_id, local, remote)
      @logger.debug("comparing #{local} to #{remote}")
      local_md5 = Digest::MD5.file(local).hexdigest
      command = <<-EOH
        $dest_file_path = [System.IO.Path]::GetFullPath('#{remote}')

        if (Test-Path $dest_file_path) {
          $crypto_prov = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
          try {
            $file = [System.IO.File]::Open($dest_file_path,
              [System.IO.Filemode]::Open, [System.IO.FileAccess]::Read)
            $guest_md5 = ([System.BitConverter]::ToString($crypto_prov.ComputeHash($file)))
            $guest_md5 = $guest_md5.Replace("-","").ToLower()
          }
          finally {
            $file.Dispose()
          }
          if ($guest_md5 -eq '#{local_md5}') {
            exit 0
          }
        }
        remove-item $dest_file_path -Force
        exit 1
      EOH
      powershell(command, shell_id)[:exitcode] == 1
    end

    def upload_to_remote(shell_id, local, remote)
      @logger.debug("Uploading '#{local}' to temp file '#{remote}'")
      base64_host_file = Base64.encode64(IO.binread(local)).gsub("\n", "")
      base64_host_file.chars.to_a.each_slice(8000 - remote.size) do |chunk|
        output = cmd("echo #{chunk.join} >> \"#{remote}\"", [], shell_id)
        raise_upload_error_if_failed(output, local, remote)
      end
    end

    def upload_directory(shell_id, local, remote)
      zipped = zip_path(local)
      return if !File.exist?(zipped)
      remote_zip = File.join(remote, File.basename(zipped))
      @logger.debug("uploading #{zipped} to #{remote_zip}")
      upload_file(zipped, remote_zip, shell_id)
      extract_zip(remote_zip, local, shell_id)
    end

    def zip_path(path)
      path.sub!(%r[/$],'')
      archive = File.join(path,File.basename(path))+'.zip'
      FileUtils.rm archive, :force=>true

      Zip::File.open(archive, 'w') do |zipfile|
        Dir["#{path}/**/**"].reject{|f|f==archive}.each do |file|
          entry = Zip::Entry.new(archive, file.sub(path+'/',''), nil, nil, nil, nil, nil, nil, ::Zip::DOSTime.new(2000))
          zipfile.add(entry,file)
        end
      end

      archive
    end

    def extract_zip(remote_zip, local, shell_id)
      @logger.debug("extracting #{remote_zip} to #{remote_zip.gsub('/','\\').gsub('.zip','')}")
      command = <<-EOH
        $shellApplication = new-object -com shell.application 
        $zip_path = remote_zip.gsub('/','\\')}

        $zipPackage = $shellApplication.NameSpace($zip_path) 
        $dest_path = "$($env:systemDrive)#{remote_zip.gsub('/','\\').gsub('.zip','')}"
        mkdir $dest_path -ErrorAction SilentlyContinue
        $destinationFolder = $shellApplication.NameSpace($dest_path) 
        $destinationFolder.CopyHere($zipPackage.Items(),0x10)
      EOH

      output = powershell(command, shell_id)
      raise_upload_error_if_failed(output, local, remote_zip)
    end

    def decode_remote_file(shell_id, local, remote)
      @logger.debug("Decoding temp file '#{remote}'")
      command = <<-EOH
        $tmp_file_path = [System.IO.Path]::GetFullPath('#{remote}')

        $dest_dir = ([System.IO.Path]::GetDirectoryName($tmp_file_path))
        New-Item -ItemType directory -Force -Path $dest_dir

        $base64_string = Get-Content $tmp_file_path
        $bytes = [System.Convert]::FromBase64String($base64_string)
        [System.IO.File]::WriteAllBytes($tmp_file_path, $bytes)
      EOH
      output = powershell(command, shell_id)
      raise_upload_error_if_failed(output, local, remote)
    end

    def raise_upload_error_if_failed(output, from, to)
      raise UploadFailed,
        :from => from,
        :to => to,
        :message => output.inspect unless output[:exitcode].zero?
    end

  end # WinRMWebService
end # WinRM
