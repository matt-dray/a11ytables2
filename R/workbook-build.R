#' Create a Workbook and Fill with Blueprint Information
#'
#' @param blueprint A list, required. A blueprint for creating a workbook with
#'   one list element for the cover, contents, notes (optional) and tables,
#'   probably generated using [append_cover], [append_contents], [append_notes]
#'   and [append_tables].
#'
#' @return A wbWorkbook object.
#'
#' @examples
#' \dontrun{
#' bp <- new_blueprint() |>
#'   append_cover() |>
#'   append_contents() |>
#'   append_notes() |>
#'   append_tables() |>
#'   append_tables()
#'
#' wb <- create_workbook(blueprint)
#' }
#'
#' @export
generate_workbook <- function(blueprint) {

  wb <- openxlsx2::wb_workbook()

  sheet_types <- list("cover", "contents", "notes", "tables")
  invisible(lapply(sheet_types, function(x) .add_sheet(blueprint, wb, x)))

  wb

}

#' Add a Sheet to the Workbook
#' @noRd
.add_sheet <- function(
    blueprint,
    wb,
    sheet_type = c("cover", "contents", "notes", "tables")
) {

  sub_blueprint <- .isolate_blueprint_sheet(blueprint, sheet_type)

  for (i in seq_along(sub_blueprint)) {

    sheet_name <- names(sub_blueprint)[i]
    sheet_content <- sub_blueprint[[i]]

    wb$add_worksheet(sheet = sheet_name)
    .insert_metadata(wb, sheet_name, sheet_content)

    if (sheet_type == "cover" & length(sub_blueprint) > 0) {
      .insert_sections(wb, sheet_name, sheet_content)
    }

    if (sheet_type != "cover" & length(sub_blueprint) > 0) {
      .insert_tables(wb, sheet_name, sheet_content)
    }

  }

}

#' Extract Blueprint by Sheet Type
#' @noRd
.isolate_blueprint_sheet <- function(blueprint, sheet_type) {
  Filter(function(x) x[["sheet_type"]] == sheet_type, blueprint)
}
