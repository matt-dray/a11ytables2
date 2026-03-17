
<!-- README.md is generated from README.Rmd. Please edit that file -->

# {a11ytables2}

<!-- badges: start -->

[![Project Status: Concept – Minimal or no implementation has been done
yet, or the repository is only intended to be a limited example, demo,
or
proof-of-concept.](https://www.repostatus.org/badges/latest/concept.svg)](https://www.repostatus.org/#concept)
<!-- badges: end -->

> \[!NOTE\] This experimental project is now archived. Please see
> [{aftables}](https://github.com/best-practice-and-impact/aftables),
> which is a renamed and updated version of the original {a11ytables}
> package.

## Purpose

Generate spreadsheet publications that follow [best-practice guidance
from the UK Government’s Analysis
Function](https://analysisfunction.civilservice.gov.uk/policy-store/releasing-statistics-in-spreadsheets/).

This package is a work-in-progress concept to experiment with new
methods for [the {a11ytables}
package](https://github.com/co-analysis/a11ytables). It may never be
fully-featured or complete.

## Install

Install from GitHub via {remotes}:

``` r
remotes::install_github("matt-dray/a11ytables2")
```

## Workflow

The basic workflow involves building a ‘blueprint’: a list object that
contains all the information needed to create an output workbook.

First, let’s define some demo data that represents a table of
statistical data to be published.

``` r
mtcars_x <- mtcars[6:10, 1:5]
mtcars_x[2, 2] <- "[note 1]"
```

Now we can create our list blueprint (`new_blueprint()`) and build it up
sheet by sheet, with specialised functions to add each sheet required by
the best-practice standards (`append_cover()`, `append_contents()`,
`append_notes()`, `append_tables()`). Each function has its own checks
and arguments.

``` r
library(a11ytables2)

blueprint <- 
  new_blueprint() |>  # initiate blueprint list
  append_cover(
    title = "I am a Test Workbook",
    sections = list(  # element names are headers, vector elements are new lines
      "First Section" = "This is some text.",
      "Second Section" = c("This is some text.", "This is some more text."),
      "Third Section" = c("This is some text.", "This is some more text", "Even more."),
      "Fourth Section" = "This is some text."
    )
  ) |> 
  append_contents(
    table = data.frame(
      Tab = c("Notes", "Table_1"),
      Description = c("The notes page.", "The first table [note 5].")
    )
  ) |>
  append_notes(
    table = data.frame(
      Note = c("[note 1]", "[note 2]"),
      Description = c("This is a note.", "This is another note")
    )
  ) |> 
  append_tables(
    sheet_name = "Table_1",  # will appear as the tab name in the workbook
    title = "The Title 1 [note 3]",
    subtitle = "The subtitle 1",
    custom = c(  # arbitrary pre-table information
      x = "A custom row",
      y = "Another custom row",
      "A third."
    ),
    source = "The source.",
    tables = mtcars_x
  ) |> 
  append_tables(  # this worksheet has two tables
    sheet_name = "Table_2",
    title = "The Title 2",
    subtitle = "The subtitle 2",
    source = "The source 2.",
    tables = list(  # provide multiple tables in list format
      "Subtable 2.1" = mtcars[1:5, 1:5],
      "Subtable 2.2" = mtcars_x
    )
  )
```

<details>

<summary>

Click to see the structure of the blueprint object.
</summary>

``` r
str(blueprint, 3)
# List of 5
#  $ Cover   :List of 6
#   ..$ sheet_type    : chr "cover"
#   ..$ title         : chr "I am a Test Workbook"
#   ..$ First Section : chr "This is some text."
#   ..$ Second Section: chr [1:2] "This is some text." "This is some more text."
#   ..$ Third Section : chr [1:3] "This is some text." "This is some more text" "Even more."
#   ..$ Fourth Section: chr "This is some text."
#  $ Contents:List of 4
#   ..$ sheet_type   : chr "contents"
#   ..$ title        : chr "Contents"
#   ..$ notes_present: chr "There are notes in this sheet."
#   ..$ table        :'data.frame': 2 obs. of  2 variables:
#   .. ..$ Tab        : chr [1:2] "Notes" "Table_1"
#   .. ..$ Description: chr [1:2] "The notes page." "The first table [note 5]."
#  $ Notes   :List of 4
#   ..$ sheet_type   : chr "notes"
#   ..$ title        : chr "Notes"
#   ..$ notes_present: chr "There are notes in this sheet."
#   ..$ table        :'data.frame': 2 obs. of  2 variables:
#   .. ..$ Note       : chr [1:2] "[note 1]" "[note 2]"
#   .. ..$ Description: chr [1:2] "This is a note." "This is another note"
#  $ Table_1 :List of 10
#   ..$ sheet_type   : chr "tables"
#   ..$ title        : chr "The Title 1 [note 3]"
#   ..$ subtitle     : chr "The subtitle 1"
#   ..$ table_count  : chr "There is one table in this sheet."
#   ..$ notes_present: chr "There are notes in this sheet."
#   ..$ x            : chr "A custom row"
#   ..$ y            : chr "Another custom row"
#   ..$ custom_3     : chr "A third."
#   ..$ source       : chr "The source."
#   ..$ tables       :'data.frame': 5 obs. of  5 variables:
#   .. ..$ mpg : num [1:5] 18.1 14.3 24.4 22.8 19.2
#   .. ..$ cyl : chr [1:5] "6" "[note 1]" "4" "4" ...
#   .. ..$ disp: num [1:5] 225 360 147 141 168
#   .. ..$ hp  : num [1:5] 105 245 62 95 123
#   .. ..$ drat: num [1:5] 2.76 3.21 3.69 3.92 3.92
#  $ Table_2 :List of 7
#   ..$ sheet_type   : chr "tables"
#   ..$ title        : chr "The Title 2"
#   ..$ subtitle     : chr "The subtitle 2"
#   ..$ table_count  : chr "There are two tables in this sheet."
#   ..$ notes_present: chr "There are notes in this sheet."
#   ..$ source       : chr "The source 2."
#   ..$ tables       :List of 2
#   .. ..$ Subtable 2.1:'data.frame':   5 obs. of  5 variables:
#   .. ..$ Subtable 2.2:'data.frame':   5 obs. of  5 variables:
```

</details>

Once we have the blueprint list, we can add spreadsheet structure and
style by converting to an {openxlsx2} wbWorkbook-class object:

``` r
wb <- generate_workbook(blueprint)
wb
# A Workbook object.
#  
# Worksheets:
#  Sheets: Cover, Contents, Notes, Table_1, Table_2 
#  Write order: 1, 2, 3, 4, 5
```

We can then apply some further manipulation to our wbWorkbook-class
object to finesse it for our needs. For example, we can set the number
format to ‘General’ for Table 1:

``` r
wb$add_numfmt(sheet = "Table_1", dims = "A10:E14", numfmt = "General")
```

And finally we can open a temporary copy of the workbook for inspection:

``` r
wb |> openxlsx2::wb_open()
```

Use `openxlsx2::wb_save()` instead to save to disk.

## Comparison to {a11ytables}

Improvements in {a11ytables2} compared to {a11ytables} include:

- [{openxlsx2}](https://janmarvin.github.io/openxlsx2/) for the
  back-end, rather than
  [{openxlsx}](https://ycphs.github.io/openxlsx/index.html)
- greater user control by building up with an `append_*()` function
  family
- greater flexibility to provide arbitrary pre-table content via
  `custom()` argument
- support for multiple tables per sheet

## Related projects

Actively-used packages include:

- [{a11ytables}](https://github.com/co-analysis/a11ytables) for R
- [{rapid.spreadsheets}](https://github.com/RAPID-ONS/rapid.spreadsheets)
  for R
- [‘gptables’](https://github.com/best-practice-and-impact/gptables) for
  Python

Another experimental project that builds on {a11ytables}:

- [{yamlsheets}](https://github.com/matt-dray/yamlsheets) for R
