# Address Cleaning

This consists of two separate components designed to help clean up messy addresses. The first component is a macro that can be run from an XLSM file. The second component are two custom Excel functions that split an address containing an Apt/Suite/Unit Number into two separate address lines. Together the two components can be used in a workflow that efficiently and somewhat accurately cleans up messy addresses in Excel.

## The Cleaning Macro

To use the first macro component you will need to make a copy of a local XLSM file that you will use to clean up your data prior to running that information through the custom functions. Start by making a new XLSM file locally on your computer, preferably in your downloads folder or desktop. The XLSM file should preferably be in the same directory where you temporarily store the files you are extracting and cleaning data. Name the file something along the lines of "AddressCleanMacro.XLSM" to avoid getting the file confused with the other files stored in this directory.

Once your macro is saved you can cut and paste the code from this repository into your macro. The logic is explained in more depth in the comments to the code. In short, the macro does the following
* Asks you to specify which column contains your addresses
* Looks for the last row in your data which contains information
* For each cell in the specified column that contains an address the macro will perform the following steps
  - Change the casing to proper case
  - Remove all periods from the name
  - Replaces all instances of "Po " with "PO " (This is case sensitive and set to capture only instances where there is a trailing space after Po in order to prevent capturing street names containing the characters "Po")
  - Replaces all instances where the proper casing function changed street number abbreviations into capital letters with lower case (1St to 1st, 2Nd to 2nd, 11Th to 11th, etc.)
  - Replace all of the following words with their proper abbreviations

Word | Abbreviation
------------ | -------------
Apartment | Apt
Avenue	|Ave
Boulevard	|Blvd
Center	|Ctr
Circle	|Cr
Court|	Ct
Drive	|Dr
East	|E
Expressway|	Expy
Heights|	Hts
Highway|	Hwy
Island|	Is
Junction|	Jct
Lake|	Lk
Lane|	Ln
Mountain|	Mtn
North|	N
Northwest|	NW
Parkway|	Pkwy
Place|	Pl
Plaza|	Plz
Ridge|	Rdg
Road|	Rd
Room|	Rm
South	|S
Southeast	|SE
Southwest|	SW
Square|	Sq
Station|	Sta
Street|	St
Suite|	Ste
Terrace|	Ter
Turnpike|	Tpke
Valley|	Vly
West|	W


### Bugs with the Macro
There are a few errors that you should be aware of with this macro. The following are noted
* Any instances of a misspelling of an word requiring abbreviated will be overlookd (Steet, Avaneu, Terace, etc.)

## The Address Split Functions

The address split functions are two custom excel functions that are saved in a Personal.XLSB file and can be used to further clean up your addresses after you run the cleaning macro. For information on how to create and save a Personal.XLSB file for using these functions see

https://support.microsoft.com/en-us/office/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566

Once a Personal.XLSB file is created and saved locally on your computer, you can use the code in this repository to create =PERSONAL.XLSB!ADDSPLIT1() and =PERSONAL.XLSB!ADDSPLIT2(). They work in the following manner:
Pass the name of the cell containing your address into each of the functions

=PERSONAL.XLSB!ADDSPLIT1(_Text_)

=PERSONAL.XLSB!ADDSPLIT2(_Text_)

Addsplit1 takes a string and will look for any Apartment, Suite, Unit numbers at the end of the address line, strip that information from the address, and return just the street address

=PERSONAL.XLSB!ADDSPLIT("123 Everywhere St Unit 1") returns "123 Everywhere St"

Addsplit2 does the exact opposite. It takes a string and looks for any Apartment, Suite, Unit numnbers and returns just the number and abbreviation without the street address

=PERSONAL.XLSB!ADDSPLIT2("123 Everywher St Unit 1") returns "Unit 1"

### Bugs
#### Ste and other abbreviations in street names
The function cannot catch are street names or abbrevations that begin with a name or abbreviation it is searching for. 
For example:
"123 Steely St" will cause Addsplit2 to return everything after Ste, since it mistakes the street name for a Ste abreviation
=PERSONAL.XLSB!ADDSPLIT1("123 Steely St") will result in "123 Steely St" being returned in the Address Line 1 cell
=PERSONAL.XLSB!ADDSPLIT2("123 Steely St") will result in "Steely St" being placed in the Address Line 2 cell

#### Floor numbers and numbers appearing before unit abbreviations
The function cannot catch floor, apartment, or unit numbers that come before a name.
For example:
123 Everywhere St 1st Fl" will cause Addplit2 to return just the abbreviation Fl
=PERSONAL.XLSB!ADDSPLIT1("123 Everywhere St 1st Fl") will result in "123 Everywhere St 1st" being returned in the Address Line 1 cell
=PERSONAL.XLSB!ADDSPLIT2("123 Everywhere St 1st Fl") will result in "Fl" being placed in the Address Line 2 cell





A | B | C
-----|-----|----|
Address | Address 1 | Address 2
This cell stores your address that needs split | This cell will contain the street address | This cell will contain the apartment, unit, suite, etc. number
123 Everywhere St | =PERSONAL.XLSB!ADDSPLIT1(A1) | =PERSONAL.XLSB!ADDSPLIT2(A1)
123 Everywhere St | 123 Everywhere St |
123 Everywhere St Apt A1 | 123 Everywhere St | Apt A1
123 Everywhere St  Ste 1 | 123 Everywhere St | Ste 1
123 Everywhere St Fl 1 | 123 Everywhere St | Fl 1
123 Everywhere St 1st Fl | 123 Everywhere St 1st | Fl
123 Steely St | 123 Steely St | Steely St
