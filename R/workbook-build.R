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

  # Insert sheets ----

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

    # Styles ----

    ## Font ----

    wb$set_base_font(
      font_size = 12,
      font_color = openxlsx2::wb_color(auto = "1"),
      font_name = "Comic Sans MS"
    )

    ## Titles ----

    .style_heading(wb, sheet_name)

    if (any(names(sheet_content) %in% "subtitle")) {
      .style_heading(wb, sheet_name, "Heading 2", "A2")
    }

    if (sheet_type == "tables" && inherits(sheet_content[["tables"]], "list")) {

      subtables <- sheet_content[["tables"]]

      n_tables <- length(subtables)
      table_widths <- lengths(subtables)

      subtable_start_row <-
        length(sheet_content[!names(sheet_content) %in% c("sheet_type", "tables")]) + 1
      subtable_start_columns <- c(1, table_widths[-length(table_widths)] + 2)

      for (i in seq_along(subtable_start_columns)) {

        subtable_title_dims <-
          openxlsx2::wb_dims(subtable_start_row, subtable_start_columns[i])

        .style_heading(wb, sheet_name, "Heading 2", subtable_title_dims)

      }

    }

    ## Cover sections ----

    if (sheet_type == "cover") {

      sections <-
        sheet_content[!names(sheet_content) %in% c("sheet_type", "title")]

      section_start_rows_i <-  # TODO: looks overengineered
        c(1, cumsum(lengths(sections)[-length(sections)]) + 2:length(sections)) + 1
      section_start_dims <- paste0("A", section_start_rows_i)

      for (i in seq_along(section_start_dims)) {
        section_title_dims <- section_start_dims[i]
        .style_heading(wb, sheet_name, "Heading 2", section_title_dims)
      }

    }

  }

}

#' Extract Blueprint by Sheet Type
#' @noRd
.isolate_blueprint_sheet <- function(blueprint, sheet_type) {
  Filter(function(x) x[["sheet_type"]] == sheet_type, blueprint)
}
