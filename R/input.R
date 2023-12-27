#' Prepare a Worksheet
#'
#' @param title Character, required.
#' @param tables List of data.frames, required. See details
#' @param ... List of character vectors, optional. Each element will become its
#'     own row in the metadata of a tables sheet.
#' @param .subtitle Character vector, optional.
#' @param .notes List of character vectors, optional.
#' @param .source Character, optional.
#'
#' @details
#' # Providing tables of data
#'
#' If the sheet has a single table, you can provide it as the only element
#' inside a list, like `list(table_1_df)`. If you are providing more than one
#' table, each should be a separate sub-list with two elements: a title and a
#' data.frame, like `list(table_1a_list, table_2a_list)` where `table_1a_list`
#' is like `list(title = "Table 1a: a table.", table = table_1a_df)`,
#' for example.
#'
#' @return A list.
#'
#' @examples \dontrun{}
#'
#' @export
prepare_sheet <- function(
    title,
    tables,
    ...,
    .subtitle = NULL,
    .source = NULL
) {

  custom_elements <- list(...)

  if (length(custom_elements) == 0) {
    custom_elements <- NULL
  }

  element_list <- c(
    list(title = title, subtitle = .subtitle),
    custom_elements,
    list(source = .source, tables = tables)
  )

  Filter(function(x) length(x) > 0, element_list)

}
