
Feature: test despegar

  @CATA
  Scenario: Test despegar PASS
    Given open the browser
    When I go to despegar home page
    Then I close the advertisement
   # And I enter the place "Buenos Aires, Ciudad de Buenos Aires, Argentina"
    And I enter the place "sd6fsd6fsdfsd6f"
    And Take a screentshot

  Scenario: Test despegar not run
    Given open the browser
    When I go to despegar home page
    Then I close the advertisement
    And I enter the place "Buenos Aires, Ciudad de Buenos Aires, Argentina"
    And Take a screentshot

  Scenario: Test despegar PASS2
    Given open the browser
    When I go to despegar home page
    Then I close the advertisement
    And I enter the place "Buenos Aires, Ciudad de Buenos Aires, Argentina"
    And Take a screentshot


  Scenario: Test despegar FAIL
    Given open the browser
    When I go to despegar home page
    And I enter the place "Buenos Aires, Ciudad de Buenos Aires, Argentina"
    And Take a screentshot