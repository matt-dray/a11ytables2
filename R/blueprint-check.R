#' Check Inputs to the Cover Sheet
#' @noRd
.check_cover_input <- function(
    blueprint, sheet_name, title, subtitle, sections
) {
  .check_title(title)
  .check_subtitle(subtitle)
  .check_sections(sections)
}

#' Check Inputs to the Contents Sheet
#' @noRd
.check_contents_input <- function(
    blueprint, sheet_name, title, subtitle, custom, table
) {
  .check_title(title)
  .check_subtitle(subtitle)
}

#' Check Inputs to the Notes Sheet
#' @noRd
.check_notes_input <- function(
    blueprint, sheet_name, title, subtitle, custom, table
) {
  .check_title(title)
  .check_subtitle(subtitle)
}

#' Check Inputs to a Tables Sheet
#' @noRd
.check_tables_input <- function(
    blueprint, sheet_name, title, subtitle, custom, source, tables
) {
  .check_title(title)
  .check_subtitle(subtitle)
}

#' Check Argument Input
#' @param x Object to check.
#' @noRd
.check_title <- function(x) {

  if (!is.character(x) || length(x) > 1) {
    stop(
      "Argument 'title' must be of class character and length 1.",
      call. = FALSE
    )
  }

}

#' Check Argument Input
#' @param x Object to check.
#' @noRd
.check_subtitle <- function(x) {

  if (!is.null(x) && !is.character(x) || length(x) > 1) {
    stop(
      "Argument 'subtitle' must be of class character and length 1.",
      call. = FALSE
    )
  }

}

#' Check Argument Input
#' @param x Object to check.
#' @noRd
.check_sections <- function(x) {

  if (!inherits(x, "list")) {
    stop("Argument 'sections' must be of class list.", call. = FALSE)
  }

  if (any(names(x) == "")) {
    stop(
      "All list elements in argument 'sections' must be named.",
      call. = FALSE
    )
  }

}
