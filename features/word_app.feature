@application
Feature: Word Application Feature
  Provide easy interface to the
  "Word.Application" COM object

  Scenario: Instantiate a word application
    Given I have Word installed
    When I ask for the word version 
    Then I should get back the word version

  Scenario: Call application multiple level member method without arguments
    Given a new Word object
    When I ask for a multiple level member
    Then I should get back the correct member value
  
  Scenario: Call application method with arguments
    Given a new Word object
    When I call a method with an argument
    Then I should get the correct method value
    
  Scenario: Call application static method with a code block
    Given nothing
    When I call the Word.open method with a block
    Then I should get the correct result from the block

  Scenario: Use application to call function with no arguments
    Given nothing
    When I pass a block to execute macro Module1.SimpleMacroNoParams in document at D:/prototypes/ruby/bdd/ca/tests/features/data/DocWithMacros.doc
    Then I should get the no arg return value true

  Scenario: Use application to call a function with multiple arguments
    Given nothing
    When I pass a block to execute macro Module1.SimpleMacroMultipleParams with arguments string1 and string2 in document at D:/prototypes/ruby/bdd/ca/tests/features/data/DocWithMacros.doc
    Then I should get the multiple args return value string1+string2
