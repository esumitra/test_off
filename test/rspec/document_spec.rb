# rspec tests for document class
# run with 'spec -c -f specdoc document_spec.rb' for BDD process
# run with 'rake rspec_tests' to run as standard process'

require File.dirname(__FILE__) + '/../../lib/msword.rb'

include MSWordUtils

describe Document do
  NewDocPath=File.expand_path(File.dirname(__FILE__) + "/../../features/data")
  NormalDocPath = NewDocPath + '/DocWithoutMacros.doc'
  MacroDocPath = NewDocPath + '/DocWithMacros.doc'
  
  it "should open an existing document" do
    doc = Document.new(NormalDocPath)
    doc.Range.Text.size.should_not equal(0)
    doc.close
  end

  it "should open a new document" do
    new_file_path = "#{NewDocPath}/#{Time.now.to_i}.doc"
    doc = Document.new
    doc.SaveAs(new_file_path)
    doc.close
    File.exists?(new_file_path).should be_true
    File.delete(new_file_path)  
  end

  it "should open an existing document with a code block" do
    result = Document.open(NormalDocPath) {|doc|
      doc.Paragraphs.Count
    }
    result.should_not equal(0)
  end
  
  it "should run a macro with no arguments in an existing document" do
    doc = Document.new(MacroDocPath)
    result = doc.macro_SimpleMacroNoParams
    result.should be_true
    doc.close
  end
  
  it "should run a macro with multiple arguments in an existing document" do
    arg1='string1'
    arg2='string2'
    doc = Document.new(MacroDocPath)
    result = doc.macro_SimpleMacroMultipleParams(arg1,arg2)
    result.should == "#{arg1}+#{arg2}"
    doc.close
  end

end
