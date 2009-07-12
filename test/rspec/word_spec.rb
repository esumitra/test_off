# rspec tests for word class
require File.dirname(__FILE__) + '/../../lib/msword.rb'

include MSWordUtils

describe Word do
  DocWithMacro = File.expand_path(File.dirname(__FILE__) + "/../../features/data/DocWithMacros.doc")
  Macro1 = "Module1.SimpleMacroNoParams"
  Macro2 = "Module1.SimpleMacroMultipleParams"
  
  it "should instantiate as an object" do
    word = Word.new
    version = word.Version
    version.should == ("11.0" || "12.0")
    word.Quit(0)
  end
  
  it "should return a correct multiple level member value" do
    word = Word.new
    templates = word.Templates.Count
    word.Quit(0)
    templates.should > 0
  end
  
  it "should return the correct value for a method with an argument" do
    word = Word.new
    points = word.CentimetersToPoints(1)
    word.Quit(0)
    points.should be_close(28.34646,0.0001)
  end
  
  it "should return the correct value for code block" do
    version = Word.open{|app| app.Version}
    version.should == ("11.0" || "12.0")
  end
  
  it "should return the correct value for a macro function with no arguments" do
    result = Word.open{|app| 
      macroDoc = app.Documents.Open(DocWithMacro)
      r = app.Run(Macro1)
      macroDoc.Close(0)
      r
    }
    result.should be_true
  end

  it "should return the correct value for a macro function with multiple arguments" do
    arg1 = 'string1'
    arg2 = 'string2'
    result = Word.open{|app| 
      macroDoc = app.Documents.Open(DocWithMacro)
      r = app.Run(Macro2,arg1,arg2)
      macroDoc.Close(0)
      r
    }
    result.should == "#{arg1}+#{arg2}"
  end

end