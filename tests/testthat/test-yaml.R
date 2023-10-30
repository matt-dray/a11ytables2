test_that("YAML is read correctly", {

  test_yaml <- tempfile(fileext = ".yaml")
  yaml::write_yaml(list(), test_yaml)

  test_csv <- tempfile(fileext = ".csv")
  yaml::write_yaml(data.frame(), test_csv)

  bp <- read_blueprint(test_yaml)

  expect_identical(class(bp), "list")
  expect_error(read_blueprint(test_csv))
  expect_error(read_blueprint("x/y/z"))
  expect_error(read_blueprint("x/y/z.yaml"))
  expect_error(read_blueprint(1))
  expect_error(read_blueprint(list()))

  unlink(test_csv)
  unlink(test_yaml)

})
