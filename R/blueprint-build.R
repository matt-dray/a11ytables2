#' Begin a New Blueprint
#'
#' @return A list.
#'
#' @examples
#' bp <- new_blueprint()
#' bp
#'
#' @export
new_blueprint <- function() {
  list()
}

#' Add a Cover Sheet to a Blueprint
#'
#' Provide the information required to create a cover sheet.
#'
#' @param blueprint List, required.
#' @param sheet_name Character vector of length 1, required.
#' @param title Character vector of length 1,  required. The title of the sheet.
#'   Will be placed in cell A1 of output spreadsheet.
#' @param subtitle Character vector of length 1, optional, default is `NULL`. A
#'   subtitle for the sheet. Will be placed in cell A2 of the output spreadsheet
#'   if provided.
#' @param sections Named list of character vectors, required. A list-element
#'   name becomes a section header and each character-vector element becomes row
#'   in that section. The content will begin from cell A2 of the output
#'   spreadsheet.
#'
#' @return A list.
#'
#' @examples
#' \dontrun{
#' bp <- new_blueprint() |>
#'   append_cover()
#'
#' bp
#' }
#'
#' @export
append_cover <- function(
    blueprint,
    sheet_name = "Cover",
    title,
    subtitle = NULL,
    sections
) {

  .check_cover_input(blueprint, sheet_name, title, subtitle, sections)

  .append_sheet_elements(
    sheet_type = "cover",
    blueprint = blueprint,
    sheet_name = sheet_name,
    title = title,
    subtitle = subtitle,
    sections = sections
  )

}

#' Add a Contents Sheet to a Blueprint
#'
#' @param blueprint List, required.
#' @param sheet_name Character vector of length 1, required.
#' @param title Character vector of length 1,  required. The title of the sheet.
#'   Will be placed in cell A1 of output spreadsheet.
#' @param subtitle Character vector of length 1, optional, default is `NULL`. A
#'   subtitle for the sheet. Will be placed in cell A2 of the output spreadsheet
#'   if provided.
#' @param custom Named character vector, optional, default is `NULL`. Each
#'   element will become its own row in the metadata of sheets with sheet type
#'   'tables'. Will be placed in the order provided between the title and the
#'   source (if provided).
#' @param table A data.frame, required. Tabular data. See details.
#'
#' @return A list.
#'
#' @examples
#' \dontrun{
#' bp <- new_blueprint() |>
#'   append_cover() |>
#'   append_contents()
#'
#' bp
#' }
#'
#' @export
append_contents <- function(
    blueprint,
    sheet_name = "Contents",
    title = "Contents",
    subtitle = NULL,
    custom = NULL,
    table
) {

  .check_contents_input(blueprint, sheet_name, title, subtitle, custom, table)

  .append_sheet_elements(
    sheet_type = "contents",
    blueprint = blueprint,
    sheet_name = sheet_name,
    title = title,
    subtitle = subtitle,
    custom = custom,
    table = table
  )

}

#' Add a Notes Sheet to a Blueprint
#'
#' @param blueprint List, required.
#' @param sheet_name Character vector of length 1, required.
#' @param title Character vector of length 1,  required. The title of the sheet.
#'   Will be placed in cell A1 of output spreadsheet.
#' @param subtitle Character vector of length 1, optional, default is `NULL`. A
#'   subtitle for the sheet. Will be placed in cell A2 of the output spreadsheet
#'   if provided.
#' @param custom Named character vector, optional, default is `NULL`. Each
#'   element will become its own row in the metadata of sheets with sheet type
#'   'tables'. Will be placed in the order provided between the title and the
#'   source (if provided).
#' @param table A data.frame, required. Tabular data. See details.
#'
#' @return A list.
#'
#' @examples
#' \dontrun{
#' bp <- new_blueprint() |>
#'   append_cover() |>
#'   append_contents() |>
#'   append_notes()
#'
#' bp
#' }
#'
#' @export
append_notes <- function(
    blueprint,
    sheet_name = "Notes",
    title = "Notes",
    subtitle = NULL,
    custom = NULL,
    table
) {

  .check_notes_input(blueprint, sheet_name, title, subtitle, custom, table)

  .append_sheet_elements(
    sheet_type = "notes",
    blueprint = blueprint,
    sheet_name = sheet_name,
    title = title,
    subtitle = subtitle,
    custom = custom,
    table = table
  )

}

#' Add a Tables Sheet to a Blueprint
#'
#' @param blueprint List, required.
#' @param sheet_name Character vector of length 1, required.
#' @param title Character vector of length 1,  required. The title of the sheet.
#'   Will be placed in cell A1 of output spreadsheet.
#' @param subtitle Character vector of length 1, optional, default is `NULL`. A
#'   subtitle for the sheet. Will be placed in cell A2 of the output spreadsheet
#'   if provided.
#' @param custom Named character vector, optional, default is `NULL`. Each
#'   element will become its own row in the metadata of sheets with sheet type
#'   'tables'. Will be placed in the order provided between the title and the
#'   source (if provided).
#' @param source Character vector of length 1, optional, default is `NULL`. The
#'   source of the data provided in the tables. Will be provided as the final
#'   row of metadata above the tables.
#' @param tables A data.frame, or a list of data.frames named with table titles,
#'     required. Tabular statistical data.
#'
#' @return A list.
#'
#' \dontrun{
#' bp <- new_blueprint() |>
#'   append_cover() |>
#'   append_contents() |>
#'   append_notes() |>
#'   append_tables() |>
#'   append_tables()
#'
#' bp
#' }
#'
#' @export
append_tables <- function(
    blueprint,
    sheet_name,
    title,
    subtitle = NULL,
    custom = NULL,
    source = NULL,
    tables
) {

  .check_tables_input(
    blueprint, sheet_name, title, subtitle, custom, source, tables
  )

  .append_sheet_elements(
    sheet_type = "tables",
    blueprint = blueprint,
    sheet_name = sheet_name,
    title = title,
    subtitle = subtitle,
    custom = custom,
    source = source,
    tables = tables
  )

}

#' Convert Input to List and Add to Blueprint
#' @noRd
.append_sheet_elements <- function(
    sheet_type = c("cover", "contents", "notes", "tables"),
    blueprint = NULL,
    sheet_name = NULL,
    title = NULL,
    subtitle = NULL,
    sections = NULL,
    custom = NULL,
    source = NULL,
    table = NULL,
    tables = NULL
) {

  if (!is.null(custom)) custom <- .name_custom_elements(custom)

  table_count <- NULL
  table_count <- if (!is.null(table)) .make_table_count_sentence(table)
  table_count <- if (!is.null(tables)) .make_table_count_sentence(tables)

  notes_present <- NULL
  has_notes <- .has_notes(title, subtitle, custom, source, table, tables)
  if (has_notes) notes_present <- "There are notes in this sheet."

  sheet_elements <- c(
    list(sheet_type = sheet_type, title = title, subtitle = subtitle),
    list(table_count = table_count, notes_present = notes_present),
    sections,
    as.list(custom),
    list(source = source, table = table, tables = tables)
  )

  sheet_elements <- Filter(function(x) length(x) > 0, sheet_elements)

  modifyList(blueprint, setNames(list(sheet_elements), sheet_name))

}

#' Name Nameless Elements in Custom Element Vector
#' @noRd
.name_custom_elements <- function(custom) {

  if (length(custom) > 0) {
    nameless_i <- which(names(custom) == "")
    custom_names <- paste0("custom_", nameless_i)
    names(custom)[nameless_i] <- custom_names
  }

  custom

}

#' Build a Sentence With the Number of Tables in a Sheet
#' @noRd
.make_table_count_sentence <- function(tables) {

  n_tables <- as.character(length(tables))
  if (is.data.frame(tables)) n_tables <- "1"

  if (n_tables == 0) return(NULL)

  n_text <- switch(
    n_tables,
    "1" = "one",   "2" = "two",   "3" = "three",
    "4" = "four",  "5" = "five",  "6" = "six",
    "7" = "seven", "8" = "eight", "9" = "nine",
    n_tables
  )

  paste0(
    "There ", if (n_tables == 1) "is " else "are ", n_text,
    " table", if (n_tables > 1) "s", " in this sheet."
  )

}

.has_notes <- function(
    title = NULL,
    subtitle = NULL,
    custom = NULL,
    source = NULL,
    table = NULL,
    tables = NULL
) {

  tables_sheet_elements  <-
    list(title, subtitle, custom, source, table, unlist(tables))

  note_regex <- "\\[[N|n]ote \\d\\]"

  elements_have_notes <- lapply(
    tables_sheet_elements,
    function(x) any(grepl(note_regex, c(x, names(x))))
  )

  any(unlist(elements_have_notes))

}
