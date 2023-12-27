#' Clean a Table Name
#' @param string Character. A table name to be cleaned.
#' @noRd
.clean_table_name <- function(string) {
  gsub("[[:punct:]]|[[:space:]]", "_", tolower(string))
}
