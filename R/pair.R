#  item1, item2, name1, name2, row vs. column.
new_Pair <- function(
  item1,
  item2,
  orientation) {
  structure(
    list(
      item1 = item1,
      item2 = item2,
      orientation = orientation
    ),
    class = "Pair"
  )
}