require File.dirname(__FILE__) + '/../../lib/msword.rb'

include MSWordUtils

#-----------------------------------------
# Scenario: Instantiate a word application
#-----------------------------------------
Given /^I have Word installed$/ do
  @wdApp = Word.new
end

When /^I ask for the word version$/ do
  @version = @wdApp.Version
end

Then /^I should get back the word version$/ do
  @version.should_not == nil
  @wdApp.Quit(0)
end
#-----------------------------------------
# Scenario: Get correct template count for word application
#-----------------------------------------
Given /^a new Word object$/ do
  @wdApp = Word.new
end

When /^I ask for a multiple level member$/ do
  @tCount = @wdApp.Templates.Count
end

Then /^I should get back the correct member value$/ do
  @tCount.should == 2
  @wdApp.Quit(0)
end
#-----------------------------------------
# Scenario: Call method with arguments
#-----------------------------------------
When /^I call a method with an argument$/ do
  @points = @wdApp.CentimetersToPoints(1)
end

Then /^I should get the correct method value$/ do
  @points.should be_close(28.34646,0.0001)
  @wdApp.Quit(0)
end
#-----------------------------------------
# Scenario: Call static method with a code block
#-----------------------------------------
Given /^nothing$/ do
end

When /^I call the Word.open method with a block$/ do
  @version = Word.open{|app| app.Version}
end

Then /^I should get the correct result from the block$/ do
  @version.should eql("11.0")
end
#-----------------------------------------
# Scenario: Call a function with no arguments in a doc
#-----------------------------------------
When /^I pass a block to execute macro (.+) in document at (.+)$/ do |macro_name, filename|
  @result = Word.open{ |app| 
    filepath = File.expand_path(File.dirname(__FILE__) + '/' + filename)
    macroDoc = app.Documents.Open(filepath)
    rs = app.Run(macro_name)
    macroDoc.Close(0)
    rs
  }
end

Then /^I should get the no arg return value (.+)$/ do |value|
  @result.to_s.should eql(value)
end
#-----------------------------------------
# Scenario: Call a function with multiple arguments in a doc
#-----------------------------------------
When /^I pass a block to execute macro (.+) with arguments (.+) and (.+) in document at (.+)$/ do |macro_name, arg1,arg2, filename|
  @result = Word.open{ |app|
    filepath = File.expand_path(File.dirname(__FILE__) + '/' + filename)
    macroDoc = app.Documents.Open(filepath)
    rs = app.Run(macro_name,arg1,arg2)
    macroDoc.Close(0)
    rs
  }
end

Then /^I should get the multiple args return value (.+)$/ do |value|
  @result.should eql(value)
end
#-----------------------------------------
# Scenario: Open an existing document
#-----------------------------------------
Given /^I open an existing document at (.+)$/ do |filename|
  @doc = Document.new(filename)
end

When /^I read the document text$/ do
  @text = @doc.Range.Text
end

Then /^I should be able to write the text$/ do
  !@text.nil?
  @doc.close
end
#-----------------------------------------
# Scenario: Execute a simple macro in a document
#-----------------------------------------
Given /^I open a document with a simple macro at (.+)$/ do |filename|
  @doc = Document.new(filename)
end

When /^I call the simple macro named (.+)$/ do |macro_name|
  puts "#{@doc.Name}!#{macro_name}"
  @fResp = @doc.Application.Run("#{@doc.Name}!#{macro_name}")
end

Then /^the simple macro should return truet$/ do
  @doc.close
  @fResp
end
