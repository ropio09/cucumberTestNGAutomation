@thisone
Feature: Note on Dashboard

  Background: 
    Given I logged into suiteCRM

  Scenario: Post a note on Dashboard
    When I post "Hello Everyone"
    Then Post should be displayed
    Then I logout from application

  Scenario: Post another note on Dashboard
    When I post "Hello all"
    Then Post should be displayed
    Then I logout from application
