#' Convert a Sheet to a Workbook
#'
#' @param sheet_content List, required. Output from [prepare_sheet].
#'
#' @return a wbWorkbook object.
#'
#' @examples \dontrun{}
#'
#' @export
add_tables_sheet <- function(sheet_content) {

  wb <- openxlsx2::wb_workbook()
  wb$add_worksheet(sheet_content[["title"]])

  .insert_metadata(wb, sheet_content)
  .insert_tables(wb, sheet_content)

  wb

}

#' Insert Rows of Metadata to a Worksheet
#' @param wb wbWorkbook object.
#' @param sheet_content List, required. Output from [prepare_sheet].
#' @noRd
.insert_metadata <- function(wb, sheet_content) {

  sheet_meta <- sheet_content[names(sheet_content) != "tables"]

  for (i in seq_along(sheet_meta)) {
    wb$add_data(x = sheet_meta[i], start_row = i)
  }

}

#' Insert One or More Tables to a Worksheet
#' @param wb wbWorkbook object.
#' @param sheet_content List, required. Output from [prepare_sheet].
#' @noRd
.insert_tables <- function(wb, sheet_content) {

  n_tables <- length(sheet_content[["tables"]])

  if (n_tables == 1) {
    .insert_table(wb, sheet_content)
  }

  if (n_tables > 1) {
    .insert_subtables(wb, sheet_content)
  }

}

#' Insert One Table to a Worksheet
#' @param wb wbWorkbook object.
#' @param sheet_content List, required. Output from [prepare_sheet].
#' @noRd
.insert_table <- function(wb, sheet_content) {

  sheet_meta <- sheet_content[names(sheet_content) != "tables"]
  table_start_row <- length(sheet_meta) + 1
  table_name <- .clean_table_name(names(sheet_content[["tables"]]))
  table_content <- sheet_content[["tables"]][[1]]

  wb$add_data_table(
    x = table_content,
    start_row = table_start_row,
    table_name = table_name,
    table_style = "none"
  )

}

#' Insert Multiple Subtables to a Worksheet
#' @param wb wbWorkbook object.
#' @param sheet_content List, required. Output from [prepare_sheet].
#' @noRd
.insert_subtables <- function(wb, sheet_content) {

  subtables <- sheet_content[["tables"]]
  table_widths <- vector("list", length(subtables))

  for (i in seq_along(subtables)) {
    table_widths[[i]] <- ncol(subtables[[i]][["table"]])
  }

  sheet_meta <- sheet_content[names(sheet_content) != "tables"]
  subtable_title_start_row <- length(sheet_meta) + 1
  subtable_start_row <- length(sheet_meta) + 2

  for (i in seq_along(subtables)) {

    subtable_content <- sheet_content[["tables"]][[i]]

    subtable_start_col <- 1

    if (i > 1) {
      subtable_start_col <- sum(unlist(table_widths[1:(i - 1)]) + 1) + 1
    }

    subtable_name <-
      .clean_table_name(names(subtable_content[["title"]]))

    # Insert subtable title
    wb$add_data(
      x = subtable_content[["title"]],
      start_row = subtable_title_start_row,
      start_col = subtable_start_col
    )

    # Insert subtable
    wb$add_data_table(
      x = subtable_content[["table"]],
      start_row = subtable_start_row,
      start_col = subtable_start_col,
      table_name = i,
      table_style = "none"
    )

  }

}
