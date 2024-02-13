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
#' @param tab_name Character vector of length 1, required.
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
    tab_name,
    title,
    subtitle = NULL,
    sections
) {

  .check_cover_input(blueprint, tab_name, title, subtitle, sections)

  .append_sheet_elements(
    sheet_type = "cover",
    blueprint = blueprint,
    title = title,
    subtitle = subtitle,
    sections = sections
  )

}

#' Add a Contents Sheet to a Blueprint
#'
#' @param blueprint List, required.
#' @param tab_name Character vector of length 1, required.
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
    tab_name = "Contents",
    title = "Contents",
    subtitle = NULL,
    custom = NULL,
    table
) {

  .check_contents_input(blueprint, tab_name, title, subtitle, custom, table)

  .append_sheet_elements(
    sheet_type = "contents",
    blueprint = blueprint,
    tab_name = tab_name,
    title = title,
    subtitle = subtitle,
    custom = custom,
    tables = table
  )

}

#' Add a Notes Sheet to a Blueprint
#'
#' @param blueprint List, required.
#' @param tab_name Character vector of length 1, required.
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
    tab_name = "Notes",
    title = "Notes",
    subtitle = NULL,
    custom = NULL,
    table
) {

  .check_notes_input(blueprint, tab_name, title, subtitle, custom, table)

  .append_sheet_elements(
    sheet_type = "notes",
    blueprint = blueprint,
    tab_name = tab_name,
    title = title,
    subtitle = subtitle,
    custom = custom,
    table = table
  )

}

#' Add a Tables Sheet to a Blueprint
#'
#' @param blueprint List, required.
#' @param tab_name Character vector of length 1, required.
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
#' @param tables List of data.frames, required. Tabular data. See details.
#'
#' @details
#' # Providing tables of data
#'
#' If the sheet has a single table, you can provide it as the only element
#' inside a list, like `list(table_1_df)`. If you are providing more than one
#' table, each should be a separate sub-list with two elements: a title and a
#' data.frame, like `list(table_1a_list, table_2a_list)` where `table_1a_list`
#' is like `list(title = "Table 1a: a table.", table = table_1a_df)`, for
#' example.
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
    tab_name,
    title,
    subtitle = NULL,
    custom = NULL,
    source = NULL,
    tables
) {

  .check_tables_input(
    blueprint, tab_name, title, subtitle, custom, source, tables
  )

  .append_sheet_elements(
    sheet_type = "tables",
    blueprint = blueprint,
    tab_name = tab_name,
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
    tab_name = NULL,
    title = NULL,
    subtitle = NULL,
    sections = NULL,
    custom = NULL,
    source = NULL,
    table = NULL,
    tables = NULL
) {

  if (!is.null(custom)) custom <- .name_custom_elements(custom)

  sheet_elements <- c(
    list(sheet_type = sheet_type, title = title, subtitle = subtitle),
    sections,
    as.list(custom),
    list(source = source, table = table, tables = tables)
  )

  sheet_elements <- Filter(function(x) length(x) > 0, sheet_elements)

  .append_to_blueprint(blueprint, tab_name, sheet_elements)

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

#' Append Content of New Sheet to Existing Blueprint
#' @noRd
.append_to_blueprint <- function(blueprint, tab_name, sheet_elements) {
  blueprint <- c(blueprint, list(sheet_elements))
  names(blueprint)[length(blueprint)] <- tab_name
  blueprint
}
