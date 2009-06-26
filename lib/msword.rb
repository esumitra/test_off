# word utilities

require 'win32ole'
require 'Win32API'

module MSWordUtils
  class WdSaveOptions
    WdDoNotSaveChanges = 0
    WdSaveChanges = -1
  end

  # Document 
  # create a new document, if path specified, opens the document at the path
  # if no path specified, creates a new document
  # options :read_only => ReadOnly, :visible => Visible  
  class Document
    def initialize(path = nil,opts = {})      
      if path.nil?        
                
      else
        @app = WIN32OLE.new('Word.Application')
        read_opt = opts[:read_only]true?:false
        @doc = @app.Documents.Open(path,true,read_opt)
      end
    end
    
    def Close(saveDoc = false)
      saveOpt = saveDoc?WdSaveChanges:WdDoNotSaveChanges
      @doc.Close(saveOpt)
      @app.Quit(WdDoNotSaveChanges)
    end
    
    def self.open(path,&block)
    end
  end
  
  # Application
  class Word
    def initialize
      @app = WIN32OLE.new('Word.Application')
    end
    
    def method_missing(msg,*args)
      @app.send(msg,*args)
    end

    def self.open()
      app = WIN32OLE.new('Word.Application')
      begin
        yield(app) if block_given?
      rescue => e
        puts e.message
      ensure
        app.Quit(0)
      end
    end
  end
end
