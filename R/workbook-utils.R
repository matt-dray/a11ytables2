#' Clean a Table Name
#' @param string Character. A table name to be cleaned.
#' @noRd
.clean_table_names <- function(string) {
  gsub("[[:punct:]]|[[:space:]]", "_", tolower(string))
}
