# Consists of a single pair, and a name.
new_Sheet <- function(
  pair,
  name,
  title) {

  structure(
    list(
      pair = pair,
      name = name,
      title = title
    ),
    class = "Sheet"
  )
}