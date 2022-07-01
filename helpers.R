

dcast_leads <- function(data, formula, ...) {
  cols <- strsplit(formula, "\\s*(\\+|~|-)\\s*")[[1]]
  cols <- c(cols, "(all)")

  if (nrow(data) != 0) {
    results <- reshape2::dcast(data,
                               as.formula(formula),
                               margins = T,
                               ...
    )
    return(results)
  } else {
    return(create_empty_df(cols))
  }
}
