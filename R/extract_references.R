replace_missing_element <- function(
  l,
  name,
  replace) {
  for (i in seq_along(l)) {
    if (is.na(l[[i]][[name]])) {
      l[[i]][[name]] <- replace
    }
  }
  l
}

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

index_reference_data <- list(
  starting_index_nonfixed = list(
    pattern = "[(starting_index)]",
    pull_names = "row_first_index",
    preset = list(
      fixed_row = FALSE,
      row_last_index = NA_real_
    )
  ),
  null_index = list(
    pattern = "",
    pull_names = NULL,
    preset = list(
      fixed_row = FALSE,
      row_first_index = NA_real_,
      row_last_index = NA_real_
    )
  ),
  starting_index_fixed = list(
    pattern = "[\\$(starting_index)]",
    pull_names = "row_first_index",
    preset = list(
      fixed_row = TRUE,
      row_last_index = NA_real_
    )
  ),
  indices_nonfixed = list(
    pattern = "[(indices)]",
    pull_names = c("row_first_index", "row_last_index"),
    preset = list(
      fixed_row = FALSE
    )
  ),
  indices_fixed = list(
    pattern = "[\\$(indices)]",
    pull_names = c("row_first_index", "row_last_index"),
    preset = list(
      fixed_row = TRUE
    )
  )
)

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
  row_last_index,
  type
) {
  structure(
    list(
      sheet = sheet,
      table = table,
      field = field,
      column = column,
      fixed_row = fixed_row,
      row_first_index = row_first_index,
      row_last_index = row_last_index,
      type = type
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
      "\\[(\\${0,1})\\(",
      names(allowed_index_regex)[i],
      "\\)\\]")
    grep_pattern <- gsub(
      find_pattern,
      replacement =
        paste0(
          "\\\\[\\1(",
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
  row_last_index = NA_character_,
  type = NA_character_)

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

create_ReferenceFormat_from_partialReferenceFormat <- function(
  partial_reference_format,
  index_type) {

  index <- index_reference_data[[index_type]]
  partial_reference_format$pattern <- sub(
    "\\[\\]",
    index$pattern,
    partial_reference_format$pattern)
  partial_reference_format$pull_names <- append(
    partial_reference_format$pull_names,
    index$pull_names
  )
  partial_reference_format$preset <- append(
    partial_reference_format$preset,
    index$preset
  )

  partial_reference_format
}

find_next_reference <- function(
  preexcel_formula,
  reference_format) {

  COMPLETE <- FALSE
  grep_pattern <- pattern_to_grep(reference_format$pattern)

  matches <- gregexpr(grep_pattern, preexcel_formula$formula)[[1]]
  if (matches[[1]] == -1) {
    return(
      list(
        preexcel_formula = preexcel_formula,
        COMPLETE = TRUE
      )
    )
  }

    x <- preexcel_formula$formula
    start <- matches[[1]]
    finish <- matches[[1]] +
      attr(matches, "match.length")[[1]] - 1
    match_str <- substr(
      x,
      start,
      finish
    )
    ref <- convert_character_to_Reference(
      match_str,
      reference_format
    )
    new_id <- length(preexcel_formula$references) + 1
    match1 <- matches[1]
    attr(match1, "match.length") <- 
      attr(matches, "match.length")[1]
    attr(match1, "index.type") <- 
      attr(matches, "index.type")[1]
    attr(match1, "useBytes") <- 
      attr(matches, "useBytes")[1]
    regmatches(
      x,
      match1
    ) <- paste0("ID:", new_id)
  preexcel_formula$references <- append(
    preexcel_formula$references,
    list(ref),
  )
  preexcel_formula$formula <- x
  return(
    list(
      preexcel_formula = preexcel_formula,
      COMPLETE = FALSE
    )
  )
}

find_references_for_format <- function(
  preexcel_formula,
  reference_format) {

  COMPLETE <- FALSE
  while (!COMPLETE) {
    result <- find_next_reference(
      preexcel_formula,
      reference_format
    )
    preexcel_formula <- result$preexcel_formula
    COMPLETE <- result$COMPLETE
  }
  preexcel_formula
}

find_references <- function(
  preexcel_formula) {
  index_types <- names(index_reference_data)
  for (i in seq_along(partial_reference_formats)) {
    for (j in seq_along(index_reference_data)) {
        reference_format <- create_ReferenceFormat_from_partialReferenceFormat(
          partial_reference_formats[[i]],
          index_types[j])
        preexcel_formula <- find_references_for_format(
          preexcel_formula,
          reference_format
        )
    }
  }
  preexcel_formula
}

extract_references.basicColumn <- function(column) {
  NULL
}

extract_references.character <- function(x) {
  preexcel_formula <- new_preExcelFormula(
    formula = x,
    references = NULL
  )
  preexcel_formula <- find_references(
  preexcel_formula
 )
  return(preexcel_formula$references)
}

extract_references.conditionalFormattingElement <- function(
  formatting) {
  result <- extract_references(formatting$formula)
  result <-
    lapply(
      result,
      function(x) {
        x$type <- "conditional_formatting"
        x
      }
    )
  result
}

extract_references.FormattingElement <- function(
  formatting) {
  result <- NULL
  return(result)
}

extract_references.formulaColumn <- function(x) {
  main_references <- extract_references(
    x$x
  )
  main_references <- replace_missing_element(
    main_references,
    "type",
    "main"
  )

  if (!inherits(x$formatting, "conditionalFormattingElement")) {
    return(main_references)
  }

  c(
    main_references,
    extract_references(x$formatting)
  )
}

extract_references.formattedFormulaField <- function(x) {
  main_references <- extract_references(
    x$x
  )
  main_references <- replace_missing_element(
    main_references,
    "type",
    "main")
  if (!inherits(x$formatting, "conditionalFormattingElement")) {
    return(main_references)
  }
  formatting_references <- extract_references(x$formatting)

  formatting_references <- replace_missing_element(
    formatting_references,
    "type",
    "conditional_formatting")
  c(
    main_references,
    formatting_references
  )
}

extract_references.Summary <- function(x) {
  result <- extract_references(x$field)
  result <- replace_missing_element(
    result,
    "type",
    "main"
  )
  result
}

extract_references.summaryRow <- function(x) {
  x <- table$summary_rows[[1]]
  do.call(
    c,
    lapply(
      x,
      extract_references
    )
  )
}

extract_references.Table <- function(
  table) {

  column_references <- do.call(
    c,
    lapply(
    table$column_set,
    extract_references))
  summary_row_references <- do.call(
    c,
    lapply(
    table$summary_rows,
    extract_references))

  output <- c(
    column_references,
    summary_row_references
  )

  output <-
    replace_missing_element(
      l = output,
      "table",
      table$name)
  output
}

extract_references.Pair <- function(
  pair) {

  outp <- c(
    extract_references(pair$item1),
    extract_references(pair$item2)
  )
  outp
}

extract_references.Sheet <- function(
  sheet) {

  output <- extract_references(
    sheet$pair
  )
  output <- replace_missing_element(
    output,
    "sheet",
    sheet$name
  )
  output
}

extract_references.sheetSet <- function(x) {
  do.call(
    c,
    lapply(
      x$sheets,
      extract_references
    )
  )
}
