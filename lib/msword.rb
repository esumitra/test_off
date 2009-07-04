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
      @app = WIN32OLE.new('Word.Application')
      read_opt = opts[:read_only]?true:false
      if path.nil?        
        @doc = @app.Documents.Add   
      else
        @doc = @app.Documents.Open(path,true,read_opt)
      end
    end
    
    def close(saveDoc = false)
      saveOpt = (saveDoc)?(WdSaveOptions::WdSaveChanges):(WdSaveOptions::WdDoNotSaveChanges)
      @doc.Close(saveOpt)
      @app.Quit(WdSaveOptions::WdDoNotSaveChanges)
    end

    def method_missing(msg,*args)
      if ((msg.id2name=~/^macro_/) == 0)
        macro_name = $'
        @app.Run(macro_name,*args)
      else
        @doc.send(msg,*args)
      end
    end
    
    def self.open(path,read_only=true)
      app = WIN32OLE.new('Word.Application')
      doc = app.Documents.Open(path,true,read_only)
      begin
        yield(doc) if block_given?
      rescue => e
        puts e.message
      ensure
        save_flag=(read_only)?(WdSaveOptions::WdDoNotSaveChanges):(WdSaveOptions::WdSaveChanges)
        doc.Close(save_flag)
        app.Quit(WdSaveOptions::WdDoNotSaveChanges)
      end      
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
        app.Quit(WdSaveOptions::WdDoNotSaveChanges)
      end
    end
  end
end
