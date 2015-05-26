These are some sample keywords, full document please refer [this](http://code.google.com/p/datadriven-for-robotium-selenium/downloads/list).

  * DFRS\_pdf\_v0.9.zip
  * DFRS\_html\_doc\_v0.9.zip


# sendKey #
Syntax:
```
 sendKey(number)
```
  * CLOSED: 0
  * DELETE: 67
  * DOWN: 20
  * ENTER: 66
  * LANDSCAPE: 0
  * LEFT: 21
  * MENU: 82
  * OPENED: 1
  * PORTRAIT: 1
  * RIGHT: 22
  * UP: 19

Example:
| sendKey | 20 |
|:--------|:---|
..to go down like you press down arrow key on phone

| sendKey | 66 |
|:--------|:---|
..to ENTER like you hit enter key on phone


---

# enterText #

Syntax:
```
 enterText(text) 
```
> Enters text into an EditText.

Example:
| enterText | this is a test |
|:----------|:---------------|
..to type "this is a test" to an edittext


---

# clickOnButton #
Syntax:
```
 clickOnButton(index/name ) 
```
> Clicks on a Button with a given index.
Example:
| clickOnButton | 1 |
|:--------------|:--|
..to click the button where index of "1"

| clickOnButton | Yes |
|:--------------|:----|
..to click the button where its name of "Yes"


---

# sleep #
Syntax:
```
 sleep(mili-seconds) 
```
> wait for mili-seconds to continue to run.
Example:
| sleep | 10000 |
|:------|:------|
..to wait for 1 second to continue to run


---

# searchText #
Syntax:
```
 searchText(text)
```
> Searches for a text string and returns true if at least one item is found with the expected text. Will automatically scroll when needed

Example:
|searchText|Don Kuminaty|
|:---------|:-----------|
..to search whole screen the text of "Don Kuminaty" and returns true if found, false if not.


---

# goBack #
Syntax:
```
 goBack
```
> To go back like BACK function on almost of Android devices
Example:
|goBack|0|
|:-----|:|
..to go back like BACK function on almost of Android devices


---

# store #
Syntax:
```
 store(<a variable>, <value>)
```
> To store a (value) to (a variable)

Example 1:
| store | varTOTAL | 10000 |
|:------|:---------|:------|
..to store "10000" to variable of "varTOTAL"

Example 2:
| store | varSUBTOTAL1 | 5000 |
|:------|:-------------|:-----|
| store | varTOTAL     | varSUBTOTAL1 |
..to store "5000" to variable of "varSUBTOTAL1", then store value of variable of "varSUBTOTAL1" to "varTOTAL". So value of "varTOTAL" should be "5000"


---

# screenshot #
Syntax:
```
 screenshot(<name>)
```
> To take a .PNG whole screen screenshot on current phone view, then give it a (name)

Example 1:
| screenshot | currentloginscreen |
|:-----------|:-------------------|
..to take a whole screen screenshot on current phone view, an image file named currentloginscreen_(archivedmiliseconds).png will be created on ../DCIM/DFRS_


---

# echo #
Syntax:
```
 echo(<a string>/ <a variable>)
```
> To write out to report/debug ddms the phrase of <a string> or a value of (a variable)

Example 1:
| echo | this is test 1 |
|:-----|:---------------|
..to write out to report/debug ddms the phrase of "this is test 1"

Example 2:
| store | varSUBTOTAL1 | 5000 |
|:------|:-------------|:-----|
| echo  | varSUBTOTAL1 | 0    |
..to write out to report/debug ddms the value of variable of "varSUBTOTAL1", the it will come "5000"


---

# for ..endfor #
Syntax:
```
 for ..endfor(<variable1>,<variable2>,<variable3>)
```
  * To go loop with (variable3)/(variable2) times, if (variable1) less than (variable3) in number
  * This **for..endfor** can be loop into each other. And **endfor** must be defined after **for**
  * (variable1),(variable2),(variable3) must be stored in NUMBER

Example 1:
|store|variable1|1|0|
|:----|:--------|:|:|
|store|variable2|5|0|
|store|variable3|1|0|
|for  |variable1|variable3|variable2|
|echo |variable1|0|0|
|endfor|0        |0|0|
..this will write out to report/debug ddms

> variable1 is 1

> variable1 is 2

> variable1 is 3

> variable1 is 4

> variable1 is 5



---

# if ..endif #

Syntax:
```
 if ..endif(<variable1>,<logic value>,<variable2>)
```
  * To go out the **if..endif** with <logic value> is WRONG, TRUE to go in
  * (logic value) should be: # (diff) , "= (equal), > (greater than), < (less than), >= (greater & equal to), <= (less & equal to)
  * There is also a **else** between **if** and **endif**, then to go to else  if (logic value) is WRONG
  * And **endif** must be defined after **if**, **if..endif** can be looped
  * (variable1),(variable2) must be stored in NUMBER or STRING

Example 1:
|store|variable1|1|0|
|:----|:--------|:|:|
|store|variable2|5|0|
|if   |variable1|#|variable2|
|echo |YES      |0|0|
|else |0        |0|0|
|echo |NO       |0|0|
|endif|0        |0|0|
..this will write out to report/debug ddms
> YES

Example 2:
|store|variable1|1|0|
|:----|:--------|:|:|
|store|variable2|5|0|
|if   |variable1|"=|variable2|
|echo |YES      |0|0|
|else |0        |0|0|
|echo |NO       |0|0|
|endif|0        |0|0|
..this will write out to report/debug ddms
> NO