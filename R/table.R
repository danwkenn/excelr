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

function() {
  col1 <- new_basicColumn(
    x = iris[1],
    formatting = NULL
  )
  col2 <- new_formulaColumn(
    x = "{first_column}",
    formatting = NULL
  )

  colset <- new_columnSet(
    first_column = col1,
    second_column = col2
  )
  title <- new_formattedField(
    x = "Test Table",
    formatting = NULL
  )
  subtitle <- new_formattedField(
    x = "Test Table subtitle",
    formatting = NULL
  )
  caption <- new_formattedField(
    x = "caption of the table",
    formatting = NULL
  )
  header <- new_columnHeaderStructure(
    new_columnHeaderSet(
      new_columnHeader(
        new_formattedField(
          x = "Top Header",
          formatting = NULL
        ),
        column_index = c(
          "first_column",
          "second_column"
        )
      )
    ),
    new_columnHeaderSet(
      new_columnHeader(
        new_formattedField(
          x = "Second Header",
          formatting = NULL
        ),
        column_index = c(
          "first_column"
        )
      ),
      new_columnHeader(
        new_formattedField(
          x = "Second Header",
          formatting = NULL
        ),
        column_index = c(
          "second_column"
        )
      )
    )
  )
  summary_rows <- list(new_summaryRow(
    new_Summary(
      field = formattedFormulaField(
        x = "SUM({first_column})",
        formatting = NULL
      ),
      column_index = "first_column"
    )
  )
  )

  column_order <- c(
    "first_column",
    "second_column"
  )

  table <- new_Table(
    name = "tab1",
    title = title,
    subtitle = subtitle,
    caption = caption,
    column_header_structure = header,
    column_set = colset,
    summary_rows = summary_rows,
    column_order = column_order
  )

  pair <- new_Pair(
  new_Pair(
    table,
    {tab <- table; tab$name <- "tab2"; tab},
    "row"
  ),
  new_Pair(
    {tab <- table; tab$name <- "tab3"; tab},
    {tab <- table; tab$name <- "tab4"; tab},
    "row"
  ),
  "column")

  sheet <- new_Sheet(
  pair,
  name = "sheet_name",
  title = "Sheet Title"
  )

  sheet_set <- new_sheetSet(
    list(
      sheet
    ),
    NULL
  )

extract_references(sheet_set)

}