@document
Feature:
  Provide easy interface to the
  "Word.Document" COM object
  
  Scenario: Open existing document
    Given an existing word document
    When I open the existing document
    Then I should get the text successfully
    
  Scenario: Create new document
    Given a random document name
    When I create a new document
    Then I should be able to save the new document correctly
    
  Scenario: Open existing document and execute code block
    Given nothing
    When I pass a code block to get the paragraph count
    Then I should get the paragraph count correctly
  
  Scenario: Open existing document and run macro
    Given an existing document with macros
    When I try to execute a simple macro
    Then I should get the correct result from the simple macro

  Scenario: Open existing document and run a macro with arguments
    Given an existing document with macros
    When I try to execute a document macro with arguments
    Then I should get the correct result from the document macro with arguments
  