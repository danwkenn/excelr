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

pattern_to_grep <- function(pattern) {
  grep_pattern <- gsub("name", allowed_name_regex, pattern)
  grep_pattern <- gsub("\\{", "\\\\{", grep_pattern)
  grep_pattern <- gsub("\\}", "\\\\}", grep_pattern)
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
    modifier = character()
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
    sheet = character(n_matches),
    table = character(n_matches),
    column = character(n_matches),
    modifier = character(n_matches)
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
  return (
    list(
      references = references,
      x = x
    )
  )
}

function() {
  pull_reference_components_with_pattern(
  x = "{sheet_name:tab2:second_column} * 10 + {sheet_name:tab2:second_column}",
  pattern = "{(name):(name):(name)}",
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
    modifier = character()
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
    modifier = character()
  )

  result <- pull_reference_components_with_pattern(
  x = x,
  pattern = "{(name):(name):(name)}",
  pull_names = c(
    "sheet", "table", "column")
  )
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
  main_references$type <- rep("main", nrow(main_references))
  if (!inherits(column$formatting, "conditionalFormattingElement")) {
    return(main_references)
  }

  rbind(
    main_references,
    extract_references(column$formatting)
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
    modifier = character()
  )
  do.call(rbind, lapply(
    table$column_set,
    extract_references))
}