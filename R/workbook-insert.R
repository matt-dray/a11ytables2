#' Insert Content to Cover Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_sections <- function(wb, sheet_name, sheet_content) {

  sheet_meta <- sheet_content[names(sheet_content) %in% c("title", "subtitle")]
  sections_start_row <- length(sheet_meta) + 1

  sheet_sections <- sheet_content[!names(sheet_content) %in% c("sheet_type", "title")]
  sections_unpacked <- unlist(c(rbind(names(sheet_sections), sheet_sections)))

  wb$add_data(
    sheet = sheet_name,
    x = sections_unpacked,
    start_row = sections_start_row
  )

}

#' Insert Rows of Metadata to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_metadata <- function(wb, sheet_name, sheet_content) {

  sheet_meta <-
    sheet_content[!names(sheet_content) %in% c("sheet_type", "tables")]

  for (i in seq_along(sheet_meta)) {
    wb$add_data(sheet = sheet_name, x = sheet_meta[[i]], start_row = i)
  }

}

#' Insert One or More Tables to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_tables <- function(wb, sheet_name, sheet_content) {

  has_table_element <- any(names(sheet_content) %in% "table")
  has_tables_element <- any(names(sheet_content) %in% "tables")

  if (has_table_element) {
    is_single_table <- TRUE
  }

  if (has_tables_element) {
    is_single_table <- is.data.frame(sheet_content[["tables"]])
  }

  if (is_single_table) {
    .insert_table(wb, sheet_name, sheet_content)
  }

  is_list_of_subtables <- inherits(sheet_content[["tables"]], "list")

  if (is_list_of_subtables) {
    .insert_subtables(wb, sheet_name, sheet_content)
  }

}

#' Insert One Table to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_table <- function(wb, sheet_name, sheet_content) {

  sheet_meta <- sheet_content[!names(sheet_content) %in% c("sheet_type", "table", "tables")]
  table_start_row <- length(sheet_meta) + 1
  table_name <- .clean_table_name(sheet_content[["title"]])
  table_content <- sheet_content[["table"]]  # cover, contents
  if (any(names(sheet_content) %in% "tables")) table_content <- sheet_content[["tables"]]

  options("openxlsx2.string_nums" = TRUE)

  wb$add_data_table(
    sheet = sheet_name,
    x = table_content,
    start_row = table_start_row,
    table_name = table_name,
    table_style = "none",
    with_filter = FALSE
  )

  options("openxlsx2.string_nums" = NULL)

  table_header_dims <-
    openxlsx2::wb_dims(table_start_row, seq_len(ncol(table_content)))

  .style_table_header(wb, table_header_dims)

}

#' Insert Multiple Subtables to a Sheet
#' @param wb A wbWorkbook object.
#' @param sheet_name Character, required.
#' @param sheet_content List, required.
#' @noRd
.insert_subtables <- function(wb, sheet_name, sheet_content) {

  subtables <- sheet_content[["tables"]]
  table_widths <- lengths(subtables)
  sheet_meta <- sheet_content[!names(sheet_content) %in% c("sheet_type", "tables")]
  subtable_title_start_row <- length(sheet_meta) + 1
  subtable_start_row <- subtable_title_start_row + 1

  for (i in seq_along(subtables)) {

    subtable_content <- sheet_content[["tables"]][i]
    subtable_start_col <- 1

    if (i > 1) {
      subtable_start_col <- sum(unlist(table_widths[1:(i - 1)]) + 1) + 1
    }

    subtable_title <- names(subtable_content)
    subtable_name <- .clean_table_name(subtable_title)
    subtable_table <- subtable_content[[1]]

    # Insert subtable title
    wb$add_data(
      sheet = sheet_name,
      x = subtable_title,
      start_row = subtable_title_start_row,
      start_col = subtable_start_col
    )

    options("openxlsx2.string_nums" = TRUE)

    # Insert subtable
    wb$add_data_table(
      sheet = sheet_name,
      x = subtable_table,
      start_row = subtable_start_row,
      start_col = subtable_start_col,
      table_name = subtable_name,
      table_style = "none",
      with_filter = FALSE
    )

    options("openxlsx2.string_nums" = NULL)

    table_header_dims <-
      openxlsx2::wb_dims(
        subtable_start_row,
        seq_len(subtable_start_col + ncol(subtable_table))
      )

    .style_table_header(wb, table_header_dims)

  }

}

#' Clean a Table Name
#' @param string Character. A table name to be cleaned.
#' @noRd
.clean_table_name <- function(string) {

  string <- tolower(string)
  string <- gsub("\\[note \\d{1,}\\]", "_", string)
  string <- trimws(string)
  gsub("[[:punct:]]|[[:space:]]", "_", string)

}
