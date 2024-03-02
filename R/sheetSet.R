# Set of sheets.
# hide sheet instructions.
# additional instructions
new_sheetSet <- function(
  sheets,
  instructions) {

  structure(
    list(
      sheets = sheets,
      instructions = instructions
    ),
    class = "sheetSet"
  )
}