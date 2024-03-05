extract_references <- function(x) {
  UseMethod("extract_references")
}

accepted_modifiers <- c("SHIFT")
modifier_regex <- paste0(
  "MODIFIER_(?:",
  paste0(
    accepted_modifiers,
    collapse = "|"
  ),
  ")\\\\((?:|(?:\\-|)\\d+(?:\\,(?:\\-|)\\d+)*\\\\)"
)
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
    "\\\\-{0,1}\\\\d+",
    "\\\\:",
    "\\\\-{0,1}\\\\d+"
  ))

# For summary row references.
allowed_summary_index_regex <- c(
  full_column = paste0(
  "SUMMARY_\\:"
  ),
  starting_index = paste0(
    "SUMMARY_\\\\d+"
  ),
  indices = paste0(
    "SUMMARY_\\\\d+",
    "\\\\:",
    "\\\\d+"
  ))


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


  for (i in seq_along(allowed_summary_index_regex)) {
    find_pattern <- paste0(
      "\\[\\(",
      names(allowed_summary_index_regex)[i],
      "\\)\\]")
    grep_pattern <- gsub(
      find_pattern,
      replacement =
        paste0(
          "\\\\[(",
          allowed_summary_index_regex[[i]],
          ")\\\\]"),
      grep_pattern
    )
  }

  grep_pattern <- gsub("\\$", "\\\\$", grep_pattern)

  grep_pattern
}

pull_reference_components_with_pattern <- function(
  x,
  pattern,
  pull_names) {

  # Coerce_pattern:
  grep_pattern <- pattern_to_grep(pattern)
  matches <- gregexpr(grep_pattern, x)
  matching_references <- regmatches(x, matches)[[1]]
  n_matches <- length(matches[[1]])
  if (matches[[1]][1] == -1) {
    n_matches <- 0
  references <- data.frame(
    reference = matching_references,
    reference_type = rep(pattern, length = n_matches),
    from = integer(),
    to = integer(),
    sheet = character(),
    table = character(),
    column = character(),
    row_first_index = integer(), row_last_index = integer()
  )
  return(
    list(
      references = references,
      x = x
    )
  )
  } else {
    references <- data.frame(
    reference = matching_references,
    reference_type = rep(pattern, length = n_matches),
    from = matches[[1]],
    to = matches[[1]] + attr(matches[[1]], "match.length") - 1,
    sheet = rep(NA_character_, n_matches),
    table = rep(NA_character_, n_matches),
    column = rep(NA_character_, n_matches),
    modifier = rep(NA_character_, n_matches)
  )
  }

  for (i in seq_along(pull_names)) {
    references[[pull_names[i]]] <- gsub(
      grep_pattern,
      paste0("\\", i),
      matching_references
    )
  }
  for (i in seq_along(matches[[1]])) {
    substr(
      x,
      matches[[1]][i],
      matches[[1]][i] + attr(
        matches[[1]], "match.length")[i] -1
      ) <- paste0(rep("X",attr(
        matches[[1]], "match.length")[i]),
        collapse = "X")
  }
  return(
    list(
      references = references,
      x = x
    )
  )
}

function() {
  pull_reference_components_with_pattern(
  x = "{sheet_name:tab2:second_column$[SUMMARY_4:10]} * 10 + {sheet_name:tab2:second_column[:]} + {first_column[3]}",
  pattern = "{(name):(name):(name)[(full_column)]}",
  pull_names = c(
    "sheet", "table", "column")
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
