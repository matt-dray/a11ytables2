#' Style Header
#' @noRd
.style_heading <- function(wb, sheet_name, style_name = "Heading 1", style_dims = "A1") {

  font_size <- 16
  if (style_name == "Heading 2") font_size <- 14

  wb$add_named_style(
    sheet = sheet_name,
    dims = style_dims,
    name = style_name,
  )

  wb$add_font(
    sheet = sheet_name,
    dims = style_dims,
    name = "Comic Sans MS",
    color = openxlsx2::wb_color(auto = "1"),
    size = font_size,
    bold = TRUE
  )

  wb$add_border(
    sheet = sheet_name,
    dims = style_dims,
    top_border = NULL,
    left_border = NULL,
    bottom_border = NULL,
    right_border = NULL
  )

}

#' Style Table Headers
#' @noRd
.style_table_header <- function(wb, dims) {

  wb$add_font(
    dims = dims,
    name = "Comic Sans MS",
    size = 12,
    color = openxlsx2::wb_colour(auto = "1"),
    bold = TRUE
  )

}
