require 'zip'

module WinRM
  class RemoteDirectory < RemoteFile

    def initialize(service, local_path, remote_path)
      @logger = Logging.logger[self]
      super(service, zip(local_path), remote_path)
      logger.debug("Upload directory: #{@local_path} -> #{@remote_path}")
    end

    def upload
      super
      extract_zip
    end

    private

    def zip(local_path)
      archive = File.join(local_path,File.basename(local_path))+'.zip'
      FileUtils.rm archive, :force=>true

      logger.debug("zipping #{local_path}/**/* to #{archive}")
      Zip::File.open(archive, 'w') do |zipfile|
        Dir["#{local_path}/**/*"].reject{|f|f==archive}.each do |file|
          entry = Zip::Entry.new(archive, file.sub(local_path+'/',''), nil, nil, nil, nil, nil, nil, ::Zip::DOSTime.new(2000))
          zipfile.add(entry,file)
        end
      end

      archive
    end

    def extract_zip
      destination = remote_path.gsub('/','\\').gsub('.zip','')
      logger.debug("extracting #{remote_path} to #{destination}")
      command = <<-EOH
        $shellApplication = new-object -com shell.application 

        $zipPackage = $shellApplication.NameSpace('#{remote_path}') 
        mkdir #{destination} -ErrorAction SilentlyContinue
        $destinationFolder = $shellApplication.NameSpace('#{destination}') 
        $destinationFolder.CopyHere($zipPackage.Items(),0x10)
      EOH

      powershell(command)
    end
  end
end