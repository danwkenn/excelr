new_formattedField <- function(
  x,
  formatting
) {
  structure(
    list(
      x = x,
      formatting = formatting
    ),
    class = c("formattedField")
  )
}

formattedFormulaField <- function(
  x,
  formatting
) {
  structure(
    list(
      x = x,
      formatting = formatting
    ),
    class = c("formattedFormulaField")
  )
}

# Column set
new_columnSet <- function(
  ...
) {
  structure(
    as.pairlist(list(
    ...
    )),
    class = "columnSet"
  )
}

# Column headers
new_columnHeader <- function(
  field,
  column_index
) {
  structure(
    list(
    field = field,
    column_index = column_index
    ),
  class = "columnHeader"
  )
}

# Column header set
new_columnHeaderSet <- function(
  ...
) {
  structure(
    list(
      ...
    ),
  class = "colHeaderSet"
  )
}

# Column Header Structure
new_columnHeaderStructure <- function(
  ...
) {
  structure(
    list(
      ...
    ),
  class = "columnHeaderStructure"
  )
}

new_Summary <- function(
  field,
  column_index
) {
  structure(
    list(
    field = field,
    column_index = column_index
    ),
    class = "Summary"
  )
}

new_summaryRow <- function(
  ...
) {
  structure(
    list(
      ...
    ),
  class = "summaryRow"
  )
}

# Table structure:
# - name
# - Title
#   - content, formatting
# - Subtitle
#   - content, formatting
# - Caption
#   - content, formatting
# - Structured column header
#   - content, formatting
# - Column headers
#   - content, formatting
# - columns/body
#     - each column has a name
#     - each column has formatting
# - summary row(s)
#   - formula, and formatting
# - column order
new_Table <- function(
  name,
  title,
  subtitle,
  caption,
  column_header_structure,
  column_set,
  summary_rows,
  column_order
) {
  structure(
    list(
      name = name,
      title = title,
      subtitle = subtitle,
      caption = caption,
      column_header_structure = column_header_structure,
      column_set = column_set,
      summary_rows = summary_rows,
      column_order = column_order
    ),
    class = "Table"
  )
}
