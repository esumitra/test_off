require 'D:/prototypes/ruby/bdd/ca/lib/msword.rb'

include MSWordUtils

NormalDocPath = "D:/prototypes/ruby/bdd/ca/tests/features/data/DocWithoutMacros.doc"
MacroDocPath= "D:/prototypes/ruby/bdd/ca/tests/features/data/DocWithMacros.doc"
NewDocPath="D:/prototypes/ruby/bdd/ca/tests/features/data"

Given /^an existing word document$/ do
  File.exists?(NormalDocPath).should be_true 
end

When /^I open the existing document$/ do
  @doc = Document.new(NormalDocPath)
end

Then /^I should get the text successfully$/ do
  @doc.Range.Text.size.should_not equal(0)
  @doc.close
end

Given /^a random document name$/ do
  @new_file_path = "#{NewDocPath}/#{Time.now.to_i}.doc"
end

When /^I create a new document$/ do
  @doc = Document.new
end

Then /^I should be able to save the new document correctly$/ do
  @doc.SaveAs(@new_file_path)
  @doc.close
  File.exists?(@new_file_path).should be_true
  File.delete(@new_file_path)
end

When /^I pass a code block to get the paragraph count$/ do
  Document.open(NormalDocPath) { |doc|
    @pars = doc.Paragraphs.Count
  }
end

Then /^I should get the paragraph count correctly$/ do
  @pars.should_not equal(0)
end

Given /^an existing document with macros$/ do
  @doc = Document.new(MacroDocPath)
end

When /^I try to execute a simple macro$/ do
  @macro_result = @doc.macro_SimpleMacroNoParams
end

Then /^I should get the correct result from the simple macro$/ do
  @macro_result.should be_true
  @doc.close
end

When /^I try to execute a document macro with arguments$/ do
  @macro_result = @doc.macro_SimpleMacroMultipleParams("string1","string2")
end

Then /^I should get the correct result from the document macro with arguments$/ do
  @macro_result.should == "string1+string2"
  @doc.close
end
