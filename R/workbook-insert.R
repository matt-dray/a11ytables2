#' Insert Content to Cover Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_sections <- function(wb, sheet_name, sheet_content) {

  sheet_sections <- sheet_content[names(sheet_content) == "sections"]

  wb$add_data(sheet = sheet_name,  x = sheet_sections, start_row = 2)

}

#' Insert Rows of Metadata to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_metadata <- function(wb, sheet_name, sheet_content) {

  sheet_meta <- sheet_content[names(sheet_content) != "tables"]

  for (i in seq_along(sheet_meta)) {
    wb$add_data(sheet = sheet_name, x = sheet_meta[i], start_row = i)
  }

}

#' Insert One or More Tables to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_tables <- function(wb, sheet_name, sheet_content) {

  n_tables <- length(sheet_content[["tables"]])

  if (n_tables == 1) {
    .insert_table(wb, sheet_name, sheet_content)
  }

  if (n_tables > 1) {
    .insert_subtables(wb, sheet_name, sheet_content)
  }

}

#' Insert One Table to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_table <- function(wb, sheet_name, sheet_content) {

  sheet_meta <- sheet_content[names(sheet_content) != "tables"]
  table_start_row <- length(sheet_meta) + 1
  table_name <- .clean_table_name(names(sheet_content[["tables"]]))
  table_content <- sheet_content[["tables"]][[1]]

  wb$add_data_table(
    sheet = sheet_name,
    x = table_content,
    start_row = table_start_row,
    table_name = table_name,
    table_style = "none"
  )

}

#' Insert Multiple Subtables to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
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
      .clean_table_names(names(subtable_content[["title"]]))

    # Insert subtable title
    wb$add_data(
      sheet = sheet_name,
      x = subtable_content[["title"]],
      start_row = subtable_title_start_row,
      start_col = subtable_start_col
    )

    # Insert subtable
    wb$add_data_table(
      sheet = sheet_name,
      x = subtable_content[["table"]],
      start_row = subtable_start_row,
      start_col = subtable_start_col,
      table_name = i,
      table_style = "none"
    )

  }

}
