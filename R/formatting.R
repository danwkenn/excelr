new_formattingElement <- function(
  attribute,
  value) {

  structure(
    list(
    attribute = attribute,
    value = value
  ), class = "formattingElement")
}

new_conditionalFormattingElement <- function(
  attribute,
  value,
  formula) {

  structure(
    list(
    attribute = attribute,
    value = value,
    formula = formula
  ),
  class = "conditionalFormattingElement")
}
