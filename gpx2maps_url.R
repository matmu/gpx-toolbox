library(xml2)
library(dplyr)
library(openxlsx)


input_file.gpx = "examples/Sicily.gpx"
output_file.xlsx = "examples/Waypoints_Sicily.xlsx"


# Read the GPX file
df.input = read_xml(input_file.gpx)


# Remove namespaces
df.input = xml_ns_strip(df.input)


# Extract waypoints
waypoints = xml_find_all(df.input, ".//wpt")


# Extract latitude, longitude, name and build link
df.output = data.frame(
    Waypoint = xml_text(xml_find_all(waypoints, "./name")),
    lat = as.numeric(xml_attr(waypoints, "lat")),
    lon = as.numeric(xml_attr(waypoints, "lon")),
    stringsAsFactors = FALSE
    ) %>%
    mutate(GoogleMaps = paste0("https://maps.google.com/?q=", lat, ",", lon, "&ll=", lat, ",", lon, "&z=12"),
           AppleMaps = paste0("https://maps.apple.com/?q=", lat, ",", lon)) %>%
    select(Waypoint, GoogleMaps, AppleMaps)


# Create a new workbook
wb = createWorkbook()


# Add a worksheet
addWorksheet(wb, "Waypoints")


# Write data to the worksheet
writeData(wb, "Waypoints", df.output, startRow = 1)


# Add hyperlinks to the second column
lapply(1:nrow(df.output), function(i) {
    writeFormula(
        wb,
        sheet = "Waypoints",
        x = paste0('HYPERLINK("', df.output$GoogleMaps[i], '")'),
        startCol = 2,
        startRow = i + 1
        )
    }) %>% invisible()


# Add hyperlinks to the third column
lapply(1:nrow(df.output), function(i) {
    writeFormula(
        wb,
        sheet = "Waypoints",
        x = paste0('HYPERLINK("', df.output$AppleMaps[i], '")'),
        startCol = 3,
        startRow = i + 1
    )
}) %>% invisible()


# Style the header row to be bold
addStyle(wb, sheet = "Waypoints", style = createStyle(textDecoration = "bold"), rows = 1, cols = 1:3, gridExpand = TRUE)


# Adjust column widths for better readability
col_widths = sapply(df.output, function(col) max(nchar(as.character(col)), na.rm = TRUE))
setColWidths(wb, "Waypoints", cols = 1:3, widths = col_widths)
# setColWidths(wb, "Waypoints", cols = 1:3, widths = "auto")


# Save the workbook
saveWorkbook(wb, output_file.xlsx, overwrite = TRUE)

