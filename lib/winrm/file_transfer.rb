require 'winrm/file_transfer/remote_file'
require 'winrm/file_transfer/remote_directory'

module WinRM
  class FileTransfer

    def self.upload(service, local_path, remote_path)
      file = nil
      if File.directory?(local_path)
        file = RemoteDirectory.new(service, local_path, remote_path)
      else
        file = RemoteFile.new(service, local_path, remote_path)
      end

      file.upload
    ensure
      file.close unless file.nil?
    end
  end
end