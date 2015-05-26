# Introduction #
This is the keyword driven automation framework which applies for Robotium & Selenium framework
DFRS is short form of datadriven for [Robotium](http://code.google.com/p/robotium/) & [Selenium](http://seleniumhq.org/) system test framework

Features
  * Keywords are designed to be familiar with non-coding testers for black-box system testing for Android phone applications
  * Keywords are developed based on Robotium (currently) & Selenium  (next phase) system test framework
  * Easy to design test case (on MS Excel 2003 or via Exel of Google doc)
  * Effectively execute test by Maven on multiple Android smart phones
  * Effectively to verify, compare & validate result by object, properties, image....
  * Short & precisely traceable Excel test report

By [KoolJ](http://facebook.com/koolj.indahouse)

Version 1.0

Copyright @ GNU license, 2011.

# What's new #

1. Team developed these keywords:
  * getCurrentButtons
  * getCurrentCheckBoxes
  * getCurrentDatePickers
  * getCurrentEditTexts
  * getCurrentGridViews
  * getCurrentImageButtons
  * getCurrentImageViews
  * getCurrentListViews
  * getCurrentProgressBars
  * getCurrentRadioButtons
  * getCurrentScrollViews
  * getCurrentSlidingDrawers
  * getCurrentSpinners
  * getCurrentTimePickers
  * getCurrentToggleButtons
  * getCurrentViews
  * getViews
  * getAllOpenedActivities
  * clearEditTextInx
  * drag
  * sendKey
  * enterText
  * sleep
  * goBack
  * goBackToActivity
  * waitForActivity
  * waitForDialogToClose
  * waitForText
  * waitForView
  * scrollToSide
  * scrollUpList
  * scrollDownList
  * scrollUp
  * scrollDown
  * store
  * echo
  * label..goto
  * for..endfor
  * clickOnButton
  * clickOnButtonInx
  * clickOnText
  * clickOnRadioButton
  * clickOnMenuItem
  * clickOnImageButton
  * clickOnImage
  * clickOnCheckBox
  * clickOnEditText
  * clickInList
  * clickOnView
  * clickLongOnView
  * clickLongOnTextAndPress
  * clickLongOnText
  * clickLongOnScreen
  * clickLongInList
  * pressMenuItem
  * pressMenuItemPR
  * setActivityOrientation
  * KJscreenshot
  * KJgetvalueText
  * KJgetproperty
  * KJcompareproperty
  * KJclick
  * KJsetText
  * KJdataread
  * store..KJgetvalueText
  * if..endif
  * if..waitForActivity
  * if..waitForDialogToClose
  * if..waitForText
  * if..waitForView
  * if..searchText
  * if..searchEditText
  * if..searchButton
  * if..isRadioButtonChecked
  * if..isRadioButtonCheckedInx
  * if..isCheckBoxChecked
  * if..isCheckBoxCheckedInx

2. Team will develop these keywords:
  * pinch & zoom
  * report in HTML format with timing each test-case
  * build the same way for Selenium frame work
  * driven data from SQlite

# Pros & Cons #
1. Pros
  * 95% saved in test effort
  * Tester can easily, understandably design system test cases in keyword-driven, not programming test by coding
  * Totally support 100% black-box functionality testing
  * Tester can effectively execute test by Maven on multiple Android smart phones & view short & precisely traceable test report
  * This is an Object Oriented testing approach
  * Free for use, not for modification, under GNU license

2. Cons
  * Just for Android phone application, not Android phone web based application
  * Less keywords at current status
  * Dependency on [Robotium](http://code.google.com/p/robotium/) & [Selenium](http://seleniumhq.org/) framework

# How do testers use this #
1. Testers must do steps on [SetupEnvironment](http://code.google.com/p/datadriven-for-robotium-selenium/downloads/list) guide first on local PC to start running test from, then

2. [Design-test-cases](http://code.google.com/p/datadriven-for-robotium-selenium/downloads/list), in EXCEL via http://docs.google.com,  test cases/suites on MS Excel (format of 2003 version), then put them in to device under test (DUT) by 2 ways:
  * Put test cases in to device(s) by manual, or
  * Download them (defined on, e.g. [URL\_batch.xls](https://docs.google.com/spreadsheet/ccc?key=0ArH023IdhMTIdGJFUlRqU0pMU3RQbFZvNXRQbG5TbXc)) via share doc feature of http://docs.google.com

3. Plug one/multi Android mobile device(s) to test (or virtual devices), create a must-precisely-folder-name on root of external hard drive: mnt/sdcard/DCIM/DFRS

4. Create project folder to store test cases & report, by:
  * In file [CONFIG.xls](https://docs.google.com/spreadsheet/ccc?key=0ArH023IdhMTIdDFOTWdBalBZSkhsWjNnZVJQbnY4d0E) , specify folder name, ex: with key of "project\_folder", i set value of "myproject"
  * Create a folder, named "myproject", inside root: mnt/sdcard/DCIM/DFRS

5. Each time when new build of system under test (SUT) project, development team will release source code to test team, then testers will run to compile DFRS to execute test cases/suites on SUT on DUT with maven as in step  11 of [SetupEnvironment](http://code.google.com/p/datadriven-for-robotium-selenium/downloads/list) guide. This is needed, because:
  * step 5 will deploy SUT to DUT
  * step 11 will deploy DFRS engine (to test SUT) to DUT, then (11.b) run the tests that are designed on DUT

6. Tester can configure to run (#4, above) by manual (a command line) or a schedule by MS Windows schedule

7. Tester can configure to run (# 4, above) by manual (a command line) or a schedule by [MS-Windows-schedule](http://www.google.com.vn/url?sa=t&rct=j&q=ms%20windows%20schedule&source=web&cd=2&ved=0CCgQFjAB&url=http%3A%2F%2Fwindows.microsoft.com%2Fen-US%2Fwindows7%2Fschedule-a-task&ei=x9HqTvTHOZGciAfLhLjBBw&usg=AFQjCNGzyvTn4StWh-41MD-F6K4tmKm5lg&sig2=190c3GwDbKhegL0aYD00UQ)

# Where are we on the world of software testing #
1. In wikipedia of software testing theory, http://en.wikipedia.org/wiki/Software_testing#System_testing, what we do is a part of automation for system test with functionality for android mobile phones

2. With CSTE certification of QAI, including the definition and theory, http://www.softwarecertifications.org/cboks/cste/te_cat1.htm, what we do are a part of automation system test with functionality for android mobile phones

3. With ISTQB certification of ISTQB institute, including the definition and theory, http://istqb.org/display/ISTQB/Foundation+Level+Documents, what we do are a part of automation system test with functionality for android mobile phones

So, you can imagine where are we...