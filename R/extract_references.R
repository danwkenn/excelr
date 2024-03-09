allowed_name_regex <- paste0(
  "\\\\w",
  "(?:_|\\\\.|[a-zA-Z0-9])*"
)

allowed_index_regex <- c(
  full_column = paste0(
  "\\:"
  ),
  starting_index = paste0(
    "\\\\-{0,1}\\\\d+"
  ),
  indices = paste0(
    "\\\\-{0,1}\\\\d+)",
    "\\\\:",
    "(\\\\-{0,1}\\\\d+"
  ))

extract_references <- function(x) {
  UseMethod("extract_references")
}

# 1. Identify new reference.
# 2. Coerce into reference format

new_Reference <- function(
  sheet,
  table,
  field,
  column,
  fixed_row,
  row_first_index,
  row_last_index
) {
  structure(
    list(
      sheet = sheet,
      table = table,
      field = field,
      column = column,
      fixed_row = fixed_row,
      row_first_index = row_first_index,
      row_last_index = row_last_index
    ),
    class = "Reference"
  )
}

new_preExcelFormula <- function(
  formula,
  references) {
  structure(
    list(
      formula = formula,
      references = references
    ),
    class = "preExcelFormula"
  )
}

pattern_to_grep <- function(pattern) {
  grep_pattern <- gsub("name", allowed_name_regex, pattern)
  grep_pattern <- gsub("\\{", "\\\\{", grep_pattern)
  grep_pattern <- gsub("\\}", "\\\\}", grep_pattern)

  for (i in seq_along(allowed_index_regex)) {
    find_pattern <- paste0(
      "\\[\\(",
      names(allowed_index_regex)[i],
      "\\)\\]")
    grep_pattern <- gsub(
      find_pattern,
      replacement =
        paste0(
          "\\\\[(",
          allowed_index_regex[[i]],
          ")\\\\]"),
      grep_pattern
    )
  }

  grep_pattern <- gsub("\\$", "\\\\$", grep_pattern)

  grep_pattern
}

new_referenceFormat <- function(
  pattern,
  pull_names,
  preset) {

  structure(
    list(
      pattern = pattern,
      pull_names = pull_names,
      preset = preset
    ),
    class = "referenceFormat"
  )
}

convert_character_to_Reference <- function(
  x,
  reference_format) {

 reference <- new_Reference(
  sheet = NA_character_,
  table = NA_character_,
  field = NA_character_,
  column = NA_character_,
  fixed_row = NA_character_,
  row_first_index = NA_character_,
  row_last_index = NA_character_)

  grep_pattern <- pattern_to_grep(reference_format$pattern)

  for (i in seq_along(reference_format$pull_names)) {
    reference[[reference_format$pull_names[i]]] <- gsub(
      grep_pattern,
      paste0("\\", i),
      x
    )
  }

  for (i in seq_along(reference_format$preset)) {
    reference[[names(
      reference_format$preset
    )[[i]]]] <- reference_format$preset[[i]]
  }
  reference
}

function() {
  x = "{S-sheet_name:T-tab2:F-body:C-second_column[4:10]}"
  reference_format <- new_referenceFormat(
  pattern = "{S-(name):T-(name):F-(name):C-(name)[(indices)]}",
  pull_names = c("sheet", "table", "field", "column", "row_first_index", "row_last_index"),
  preset = list(
    fixed_row = FALSE
  )
  )
  convert_character_to_Reference(
    x,
    reference_format
  )
}

extract_references.basicColumn <- function(column) {
  references <- data.frame(
    reference = character(),
    reference_type = character(),
    from = integer(),
    to = integer(),
    sheet = character(),
    table = character(),
    column = character(),
    row_first_index = integer(), row_last_index = integer()
  )
  references
}

extract_references.character <- function(x) {
  references <- data.frame(
    reference = character(),
    from = integer(),
    to = integer(),
    sheet = character(),
    table = character(),
    column = character(),
    row_first_index = integer(),
    row_last_index = integer(),
    row_fixed = logical()
  )

  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name):(name)}",
  pull_names = c(
    "sheet", "table", "column")
  )
  result$references$row_first_index <- rep(1L, nrow(result$references))
  result$references$row_last_index <- rep(NA_integer_, nrow(result$references))
  result$references$row_fixed <- rep(FALSE, nrow(result$references))
  references <- rbind(
    references,
    result$references
  )
  x <- result$x
  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name)}",
  pull_names = c(
    "table", "column")
  )
  result$references$row_first_index <- rep(1L, nrow(result$references))
  result$references$row_last_index <- rep(NA_integer_, nrow(result$references))
  result$references$row_fixed <- rep(FALSE, nrow(result$references))
  references <- rbind(
    references,
    result$references
  )

  x <- result$x
  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name)}",
  pull_names = c(
    "column")
  )
  result$references$row_first_index <- rep(1L, nrow(result$references))
  result$references$row_last_index <- rep(NA_integer_, nrow(result$references))
  result$references$row_fixed <- rep(FALSE, nrow(result$references))
  references <- rbind(
    references,
    result$references
  )
  x <- result$x

  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name):(name)$[(starting_index)]}",
  pull_names = c(
    "sheet", "table", "column", "row_first_index")
  )
  result$references$row_first_index <- as.integer(
    result$references$row_first_index
  )
  result$references$row_last_index <- rep(NA_integer_, nrow(result$references))
  result$references$row_fixed <- rep(TRUE, nrow(result$references))
  references <- rbind(
    references,
    result$references
  )
  x <- result$x

  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name):(name)[(starting_index)]}",
  pull_names = c(
    "sheet", "table", "column", "row_first_index")
  )
  result$references$row_first_index <- as.integer(
    result$references$row_first_index
  )
  result$references$row_last_index <- rep(NA_integer_, nrow(result$references))
  result$references$row_fixed <- rep(FALSE, nrow(result$references))
  references <- rbind(
    references,
    result$references
  )
  x <- result$x

  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name):(name)$[(indices)]}",
  pull_names = c(
    "sheet", "table", "column", "indices")
  )
  result$references$row_first_index <- as.integer(
    sub("^(\\-{0,1}\\d+)\\:(\\-{0,1}\\d+)$", "\\1", result$references$indices)
  )
  result$references$row_last_index <- as.integer(
    sub("^(\\-{0,1}\\d+)\\:(\\-{0,1}\\d+)$", "\\2", result$references$indices)
  )
  result$references$row_fixed <- rep(TRUE, nrow(result$references))
  result$references$indices <- NULL
  references <- rbind(
    references,
    result$references
  )
  x <- result$x

  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name):(name)[(indices)]}",
  pull_names = c(
    "sheet", "table", "column", "indices")
  )
  result$references$row_first_index <- as.integer(
    sub("^(\\-{0,1}\\d+)\\:(\\-{0,1}\\d+)$", "\\1", result$references$indices)
  )
  result$references$row_last_index <- as.integer(
    sub("^(\\-{0,1}\\d+)\\:(\\-{0,1}\\d+)$", "\\2", result$references$indices)
  )
  result$references$row_fixed <- rep(FALSE, nrow(result$references))
  result$references$indices <- NULL
  references <- rbind(
    references,
    result$references
  )
  x <- result$x

  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name):(name)[(full_column)]}",
  pull_names = c(
    "sheet", "table", "column")
  )
  result$references$row_first_index <- rep(0, nrow(result$references))
  result$references$row_last_index <- rep("N", nrow(result$references))
  result$references$row_fixed <- rep(FALSE, nrow(result$references))
  result$references$indices <- NULL
  references <- rbind(
    references,
    result$references
  )
  x <- result$x

  return(references)
}

extract_references.conditionalFormattingElement <- function(
  formatting) {
  result <- extract_references(formatting$formula)
  result$type <-
    rep("conditional_formatting", nrow(result))
}

extract_references.FormattingElement <- function(
  formatting) {
  result <- extract_references("")
  result$type <- character()
  return(result)
}

extract_references.formulaColumn <- function(column) {
  main_references <- extract_references(
    column$x
  )
  main_references$type <- rep("column:main", nrow(main_references))
  if (!inherits(column$formatting, "conditionalFormattingElement")) {
    return(main_references)
  }

  rbind(
    main_references,
    extract_references(column$formatting)
  )
}

extract_references.formattedFormulaField <- function(x) {
  main_references <- extract_references(
    x$x
  )
  main_references$type <- rep("main", nrow(main_references))
  if (!inherits(x$formatting, "conditionalFormattingElement")) {
    return(main_references)
  }
  formatting_references <- extract_references(x$formatting)
  formatting_references$type <- rep("formatting", nrow(formatting_references))
  rbind(
    main_references,
    formatting_references
  )
}

extract_references.Summary <- function(x) {
  extract_references(x$field)
}

extract_references.summaryRow <- function(x) {
  x <- table$summary_rows[[1]]
  do.call(
    rbind,
    lapply(
      x,
      extract_references
    )
  )
}

extract_references.Table <- function(
  table) {

  references <- data.frame(
    reference = character(),
    from = integer(),
    to = integer(),
    sheet = character(),
    table = character(),
    column = character(),
    row_first_index = integer(), row_last_index = integer()
  )
  column_references <- do.call(rbind, lapply(
    table$column_set,
    extract_references))
  summary_row_references <- do.call(
    rbind,
    lapply(
    table$summary_rows,
    extract_references))
  summary_row_references$type <- paste0(
    "summary_row:",
    summary_row_references$type
  )
  output <- rbind(
    column_references,
    summary_row_references
  )
  output$table[is.na(output$table)] <- table$name
  output
}

extract_references.Pair <- function(
  pair) {

  references <- data.frame(
    reference = character(),
    from = integer(),
    to = integer(),
    sheet = character(),
    table = character(),
    column = character(),
    row_first_index = integer(), row_last_index = integer()
  )

  outp <- list(
    extract_references(pair$item1),
    extract_references(pair$item2)
  )

  do.call(rbind, outp)
}

extract_references.Sheet <- function(
  sheet) {

  output <- extract_references(
    sheet$pair
  )
  output$sheet[is.na(output$sheet)] <- sheet$name
  output
}

extract_references.sheetSet <- function(x) {
  do.call(
    rbind,
    lapply(
      x$sheets,
      extract_references
    )
  )
}
