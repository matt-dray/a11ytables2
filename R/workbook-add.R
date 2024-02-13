
#' Add a Cover Sheet
#'
#' Add to the workbook a cover sheet. This will be the first sheet in the
#' workbook.
#'
#' @param sheet_content List, required. Output from [prepare_sheet].
#'
#' @return A wbWorkbook object.
#'
#' @examples \dontrun{}
#'
#' @export
add_cover_sheet <- function(wb, sheet_content) {

  wb$add_worksheet(sheet_content[["title"]])

  .insert_cover(wb, sheet_content)

  wb

}

#' Add a Sheet with Tabular Content
#'
#' Add to the workbook a sheet with data in tabular form. This includes sheet
#' types 'contents', 'notes' and 'tables'. Contents and notes sheets will be
#' the second and third sheet in the workbook after the sheet with type 'cover'
#' and there can only be one of each.
#'
#' @param sheet_content List, required. Output from [prepare_sheet].
#'
#' @return A wbWorkbook object.
#'
#' @examples \dontrun{}
#'
#' @export
add_tabular_sheet <- function(wb, sheet_name, sheet_content) {

  wb$add_worksheet(sheet_content[["title"]])

  .insert_metadata(wb, sheet_name, sheet_content)
  .insert_tables(wb, sheet_name, sheet_content)

  wb

}
