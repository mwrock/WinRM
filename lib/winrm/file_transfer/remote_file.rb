module WinRM
  class RemoteFile

    attr_reader :local_path
    attr_reader :remote_path

    def initialize(service, local_path, remote_path)
      @logger = Logging.logger[self]
      @service = service
      @shell = service.open_shell
      @local_path = local_path
      @remote_path = full_remote_path(local_path, remote_path)
    ensure
      if !shell.nil?
        ObjectSpace.define_finalizer( self, self.class.close(shell, service) )
      end
    end

    def upload
      if should_upload_file?
        size = upload_to_remote
        decode_remote_file
      else
        size = 0
        logger.debug("Files are equal. Not copying #{local_path} to #{remote_path}")
      end
      size
    end
    
    def close
      service.close_shell(shell) unless shell.nil?
    end

    protected

    attr_reader :logger
    attr_reader :service
    attr_reader :shell

    def self.close(shell_id, service)
      proc { service.close_shell(shell_id) }
    end

    def full_remote_path(local_path, remote_path)
      base_file_name = File.basename(local_path)
      if File.basename(remote_path) != base_file_name
        remote_path = File.join(remote_path, base_file_name)
      end

      command = <<-EOH
        $dest_file_path = [System.IO.Path]::GetFullPath('#{remote_path}')

        if (!(Test-Path $dest_file_path)) {
          $dest_dir = ([System.IO.Path]::GetDirectoryName($dest_file_path))
          New-Item -ItemType directory -Force -Path $dest_dir | Out-Null
        }

        $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath("#{remote_path}")
      EOH

      powershell(command)
    end

    def should_upload_file?
      logger.debug("comparing #{local_path} to #{remote_path}")
      local_md5 = Digest::MD5.file(local_path).hexdigest
      command = <<-EOH
        $dest_file_path = [System.IO.Path]::GetFullPath('#{remote_path}')

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
            return $false
          }
        }
        if(Test-Path $dest_file_path){remove-item $dest_file_path -Force}
        return $true
      EOH
      out = powershell(command)
      out == 'True'
    end

    def upload_to_remote
      logger.debug("Uploading '#{local_path}' to temp file '#{remote_path}'")
      base64_host_file = Base64.encode64(IO.binread(local_path)).gsub("\n", "")
      base64_array = base64_host_file.chars.to_a
      base64_array.each_slice(8000 - remote_path.size) do |chunk|
        cmd("echo #{chunk.join} >> \"#{remote_path}\"")
      end
      base64_array.length
    end

    def decode_remote_file
      logger.debug("Decoding '#{remote_path}'")
      command = <<-EOH
        $base64_string = Get-Content '#{remote_path}'
        $bytes = [System.Convert]::FromBase64String($base64_string)
        [System.IO.File]::WriteAllBytes('#{remote_path}', $bytes)
      EOH
      powershell(command)
    end

    def powershell(script)
      script = "$ProgressPreference='SilentlyContinue';" + script
      script = script.encode('UTF-16LE', 'UTF-8')
      script = Base64.strict_encode64(script)
      cmd("powershell", ['-encodedCommand', script])
    end

    def cmd(command, arguments = [])
      command_output = nil
      out_stream = []
      err_stream = []
      service.run_command(shell, command, arguments) do |command_id|
        command_output = service.get_command_output(shell, command_id) do |stdout, stderr|
          out_stream << stdout if stdout
          err_stream << stderr if stderr
        end
      end

      if !command_output[:exitcode].zero? or !err_stream.empty?
        raise UploadFailed,
          :from => local_path,
          :to => remote_path,
          :message => command_output.inspect
      end
      out_stream.join.chomp
    end
  end
end