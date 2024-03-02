# Basic
new_basicColumn <- function(
  x,
  formatting
) {

  structure(
    list(
      x = x,
      formatting = formatting
    ),
    class = c("basicColumn")
  )
}

# Formula
new_formulaColumn <- function(
  x,
  formatting
) {

  structure(
    list(
      x = x,
      formatting = formatting
    ),
    class = c("formulaColumn")
  )
}
