#' Read a Blueprint File
#'
#' Read to a list a 'blueprint': a YAML file that encodes the content and
#' structure of a workbook compliant with Analysis Function best-practice
#' spreadsheet guidelines.
#'
#' @param yaml_path Character. A path to a YAML file that contains a valid
#'      blueprint for a compliant workbook. See details.
#'
#' @details The YAML file must declare certain key-value pairs to be considered
#'     valid.
#'
#' @return A nested list. The exact content and structure will depend on the
#'     supplied blueprint YAML file.
#'
#' @examples
#' # Path to an example blueprint YAML file
#' example_path <- system.file("extdata", "example.yaml", package = "a11ytables2")
#'
#' # Read in the blueprint YAML file to a list
#' blueprint_list <- read_blueprint(example_path)
#'
#' blueprint_list
#'
#' @export
read_blueprint <- function(yaml_path) {

  is_char <- inherits(yaml_path, "character")

  is_char_msg <- c(
    "`yaml_path` must be of class 'character'.",
    x = "You provided an object of class '{class(yaml_path)}'.",
    i = "Use an existing file path with extension '.yaml'."
  )

  if (!is_char) {
    cli::cli_abort(is_char_msg)
  }

  is_yaml <- tools::file_ext(yaml_path) %in% c("yaml", "yml")

  is_yaml_msg <- c(
    "`yaml_path` must be an existing file path with extension '.yaml'.",
    x = "You provided a file path with extension '.{tools::file_ext(path)}'."
  )

  if (!is_yaml) {
    cli::cli_abort(is_yaml_msg)
  }

  yaml_path <- fs::as_fs_path(yaml_path)

  yaml_exists <- fs::is_file(yaml_path)

  yaml_exists_msg <-       c(
    "Can't find that file path.",
    x = "You provided: {path}",
    i = "`yaml_path` must be an existing file path with extension '.yaml'."
  )

  if (!yaml_exists) {
    cli::cli_abort(yaml_exists_msg)
  }

  yaml::read_yaml(yaml_path)

}

#' Convert a Blueprint to a Workbook
#'
#' Accept a 'blueprint' list, possibly created with [read_blueprint], and
#' convert its content and structure to an 'openxlsx2' wbWorkbook-class object.
#'
#' @param blueprint List.
#'
#' @return An 'openxlsx' wbWorkbook- and R6-class object.
#'
#' @examples
#' # Path to an example blueprint YAML file
#' example_path <- system.file("extdata", "example.yaml", package = "a11ytables2")
#'
#' # Read in the blueprint YAML file to a list
#' blueprint_list <- read_blueprint(example_path)
#'
#' # Convert list to wbWorkbook-class object
#' workbook <- convert_blueprint_to_workbook(blueprint_list)
#'
#' workbook
#'
#' # openxlsx2::wb_open(workbook)  # to open temp version in spreadsheet editor
#'
#' @export
convert_blueprint_to_workbook <- function(blueprint) {

  is_list <- inherits(blueprint, "list")

  is_list_msg <- c(
    "`blueprint` must be of class 'list'.",
    x = "You provided an object of class '{class(blueprint)}'."
  )

  if (!is_list) {
    cli::cli_abort(is_list_msg)
  }

  # Extract data
  tab_titles <- names(blueprint)
  sheet_titles <- lapply(blueprint, function(tab) tab[["sheet_title"]])

  # Create new workbook
  wb <- openxlsx2::wb_workbook()

  # Add all tabs
  for (tab_title in tab_titles) {
    wb <- openxlsx2::wb_add_worksheet(wb, tab_title)
  }

  # Add all sheet titles
  for (tab_title in names(sheet_titles)) {
    wb <- openxlsx2::wb_add_data(
      wb,
      tab_title,
      sheet_titles[[tab_title]],
      start_row = 1,
      start_col = 1
    )
  }

  for (tab in tab_titles) {

    if (tab == "cover") {

      cover_vector <-
        blueprint[["cover"]][-1] |>

        stack() |>
        `[`(_, c("ind", "values")) |>
        t() |>
        as.vector()

      wb <-
        openxlsx2::wb_add_data(
          wb,
          sheet = "cover",
          x = cover_vector,
          start_row = 2
        )

    }

    if (tab == "contents") {

      contents_table <-
        sheet_titles |>
        stack() |>
        `[`(_, c("ind", "values"))

      names(contents_table) <- c("Tab title", "Sheet title")

      wb <-
        openxlsx2::wb_add_data_table(
          wb,
          sheet = "contents",
          x = contents_table,
          start_row = 2
        )

    }

  }

  wb

}
