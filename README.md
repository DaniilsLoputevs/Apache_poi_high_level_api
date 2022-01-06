# Apache poi high level API
#### Last ver: 1.2.1

## How to use
* easy example
  - [go to code](src/test/java/xlsx/Examples.java)
  - [go to result xlsx file](src/test/java/xlsxfiles/example_result_easy.xlsx)

![Img](github/img_xlsx_example_easy.png?raw=true "Output example easy")


* complex example 
  - [go to code](src/test/java/xlsx/Examples.java)
  - [go to result xlsx file](src/test/java/xlsxfiles/example_result_complex.xlsx)

![Img](github/img_xlsx_example_complex.png?raw=true "Output example hard")

## path notes
#### 1.2.1
* all side alignment
* bug fixes
#### 1.2
* change public api style
* docs
* refactoring CellGroupSelector
* refactoring packages
* IndexedColours support for cell style
* data block of CompletableFuture
#### 1.1
* CellGroupSelector - easy way to create merge region and other operations for selected cells
#### 1.0
* declarative xlsx, no reflection.


#### Dependencies
* apache poi (5.1.0)
* lombok (1.18.22)
