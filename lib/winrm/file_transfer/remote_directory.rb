require 'zip'

module WinRM
  class RemoteDirectory < RemoteFile

    def initialize(service, local_path, remote_path)
      super
      @local_path = zip
      @remote_path = File.join(@remote_path, File.basename(@local_path))
      puts "************Upload directory: #{local_path} -> #{remote_path}"
      logger.debug("Upload directory: #{local_path} -> #{remote_path}")
    end

    def upload
      super
      extract_zip
    end

    private

    def zip
      path.sub!(%r[/$],'')
      archive = File.join(path,File.basename(local_path))+'.zip'
      FileUtils.rm archive, :force=>true

      Zip::File.open(archive, 'w') do |zipfile|
        Dir["#{local_path}/**/**"].reject{|f|f==archive}.each do |file|
          entry = Zip::Entry.new(archive, file.sub(local_path+'/',''), nil, nil, nil, nil, nil, nil, ::Zip::DOSTime.new(2000))
          zipfile.add(entry,file)
        end
      end

      archive
    end

    def extract_zip
      logger.debug("extracting #{remote_path} to #{remote_path.gsub('/','\\').gsub('.zip','')}")
      command = <<-EOH
        $shellApplication = new-object -com shell.application 
        $zip_path = remote_path.gsub('/','\\')}

        $zipPackage = $shellApplication.NameSpace($zip_path) 
        $dest_path = "$($env:systemDrive)#{remote_path.gsub('/','\\').gsub('.zip','')}"
        mkdir $dest_path -ErrorAction SilentlyContinue
        $destinationFolder = $shellApplication.NameSpace($dest_path) 
        $destinationFolder.CopyHere($zipPackage.Items(),0x10)
      EOH

      powershell(command)
    end
  end
end