loadDigestDataFromXL_Severity <- function(xlsPath, sheets, labelCol="B", positionCol="C", dataCol="D") {
  # Loads digest data from Excel. 
  #
  # The data for fire severity samples are in separate worksheets.
  # Unlike the similar fire frequency data files, the sheet names do not
  # always match sample label, so instead this script examines the contents
  # of the label column and takes the last non-blank entry as the name
  # of the sample (earlier entries can be header label or other comments).
  # 
  # The data are formatted in blocks with header rows and blank lines
  # so this script has to match things up.
  #
  # The value returned is a single data.frame with cols for sample label
  # and data value (both as character).
  
  require(xlsx)
  require(stringr)
  
  gc()
  
  wb <- loadWorkbook(xlsPath)
  wss <- getSheets(wb)[sheets]
  
  getColValues <- function(ws, col, rowIndices) {
    if (missing(rowIndices)) 
      rows <- getRows(ws)
    else
      rows <- getRows(ws, rowIndices)
    
    # sometimes we end up with NAs for blank cells - no idea why
    vals <- sapply(getCells(rows, col), getCellValue)
    vals[ is.na(vals) ] <- ""
    
    vals
  }
  
  cellNamesToRowIndex <- function(cellNames) {
    # names have the form rownum.colnum
    rows <- sapply( str_split(cellNames, "\\."), function(parts) parts[1] )
    as.integer(rows)
  }
  
  lastLabel <- function(labels) {
    ts <- str_trim(labels)
    ii <- which(str_length(ts) > 0)
    ilast <- max(ii)
    labels[ilast]
  }
  
  # browser()
  
  out <- NULL
  for (isheet in 1:length(wss)) {
    
    cat("sheet", isheet, ": ")
    
    labelCol <- xlColLabelToIndex(labelCol)
    labels <- getColValues(wss[[isheet]], labelCol)
    
    # Assume that the last non-blank label col value is the
    # sample label
    sampleLabel <- lastLabel(labels)
    cat(sampleLabel, "\n")
    
    # The rows that we want are identified by having a label
    # identical to the sample label (e.g. "K1F1")
    ii <- labels == sampleLabel
    
    # We need to get the excel row numbers from the label vector names
    # because of some cols having or not having header data
    ii.rows <- cellNamesToRowIndex( names(ii)[ii] )
    
    positionCol <- xlColLabelToIndex(positionCol)
    positionVals <- getColValues(wss[[isheet]], positionCol, ii.rows)
    
    dataCol <- xlColLabelToIndex(dataCol)
    dataVals <- getColValues(wss[[isheet]], dataCol, ii.rows)
    
    if (is.null(out)) 
      out <- data.frame(label=labels[ii], position=positionVals, x=dataVals, 
                        stringsAsFactors=FALSE)
    else
      out <- rbind(out, data.frame(label=labels[ii], position=positionVals, x=dataVals, 
                                   stringsAsFactors=FALSE))
    
  }
  
  out
}
