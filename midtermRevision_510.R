# Libraries
library(tidyquant)
library(tidyverse)
library(timetk)
library(openxlsx)

# Draws stock data from three individual industries (Tech, Heath, & Retail) 
# Calculates portfolio value, portfolio cost, and positioning 
# time_series_analysis graphs of the stocks by industry
# snapshots and saves the data into an excel file, to analyze 3 industries over the course of three months 
options(warn=-1)

# Table 1 
# manipulation 
Tech_data_tbl <- c("AAPL","GOOG","NFLX","NVDA") %>%
  tq_get(from = "2021-07-30", to = "2021-10-30")
tech_stocks = data.frame(Tech_data_tbl)

# Calculates portfolio value, portfolio cost, and positioning 
tech_stocks$Units <- round(runif(1, 1, 2))
tech_stocks$Price_Bought <- (runif(1, 400, 900))
tech_stocks$Portfolio_Cost <- tech_stocks$Price_Bought * tech_stocks$Units
tech_stocks$Portfolio_Price <- tech_stocks$Units * tech_stocks$adjusted
tech_stocks$Portfolio_Revenue <- tech_stocks$Portfolio_Price - tech_stocks$Portfolio_Cost
tech_stocks

# Investment portfolio calculations 
port_Cost_T = sum(tech_stocks$Portfolio_Cost)
port_Value_T = sum(tech_stocks$Portfolio_Price)
# Tech_position = sum(tech_stocks$Portfolio_Revenue)
Tech_position = port_Value_T - port_Cost_T
Tech_position

# Tech_position - bullish 
cat("Value of Tech stocks portfolio: $", port_Value_T)
cat("Cost of Tech stocks portfolio: $", port_Cost_T)
cat("Position on Tech stocks portfolio: $", Tech_position)

# tech_plot_table - for graph 
tech_plot_table <- select(tech_stocks, symbol, date, adjusted, Portfolio_Revenue)

#graph
tech_plot <- tech_plot_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, 
                   .facet_ncol = 2, 
                   .facet_scales = "free",
                   .interactive = FALSE)
tech_plot

# Archive the data into a excel file to create snapshots in time of your portfolio, and stock performance
#create workbook
wb <- createWorkbook()

# add Worksheets
addWorksheet(wb, sheetName = "Stock_Anysis_Tech", gridLines = FALSE)

# add plot
print(tech_plot)
insertPlot(wb, 1, width = 5, height = 3.5, fileType = "png", units = "in")

# save the workbook
saveWorkbook(wb, "005_excel_workbook.xlsx", overwrite = TRUE)




# Health Industry
# Table 2 
Health_data_tbl <- c("TDOC","PFE","MRNA","CVS") %>%
  tq_get(from = "2021-07-30", to = "2021-10-30")
health_stocks = data.frame(Health_data_tbl)

# Calculates portfolio value, portfolio cost, and positioning 
health_stocks$Units <- round(runif(1, 2, 5))
health_stocks$Price_Bought <- (runif(1, 100, 130))
health_stocks$Portfolio_Cost <-  health_stocks$Price_Bought *  health_stocks$Units
health_stocks$Portfolio_Price <-  health_stocks$Units *  health_stocks$adjusted
health_stocks$Portfolio_Revenue <-  health_stocks$Portfolio_Price -  health_stocks$Portfolio_Cost
health_stocks

# Investment portfolio calculations 
port_Cost_H = sum(health_stocks$Portfolio_Cost)
port_Value_H = sum(health_stocks$Portfolio_Price)

# Health_position = sum(tech_stocks$Portfolio_Revenue)
Health_position = port_Value_H - port_Cost_H
Health_position

# Health_position - bullish 
cat("Value of Health stocks portfolio: $", port_Value_H)
cat("Cost of Health stocks portfolio: $", port_Cost_H)
cat("Position on Health stocks portfolio: $", Health_position)

# health_plot_table - for graph 
health_plot_table <- select(health_stocks, symbol, date, adjusted, Portfolio_Revenue)

#graph
health_plot <- health_plot_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, 
                   .facet_ncol = 2, 
                   .facet_scales = "free",
                   .interactive = FALSE)
health_plot

# add worksheet
addWorksheet(wb, sheetName = "Stock_Anysis_Health", gridLines = FALSE)

# add plot
print(health_plot)
insertPlot(wb, 2, width = 5, height = 3.5, fileType = "png", units = "in")

# save the workbook
saveWorkbook(wb, "005_excel_workbook.xlsx", overwrite = TRUE)




# Retail Indsutry 
# Table 3 
Retail_data_tbl <- c("LULU","TGT","WMT","CRI") %>%
  tq_get(from = "2021-07-30", to = "2021-10-30")
retail_stocks = data.frame(Retail_data_tbl)

# Calculates portfolio value, portfolio cost, and positioning 
retail_stocks$Units <- round(runif(1, 2, 10))
retail_stocks$Price_Bought <- (runif(1, 100, 195))
retail_stocks$Portfolio_Cost <- retail_stocks$Price_Bought *  retail_stocks$Units
retail_stocks$Portfolio_Price <- retail_stocks$Units *  retail_stocks$adjusted
retail_stocks$Portfolio_Revenue <- retail_stocks$Portfolio_Price -  retail_stocks$Portfolio_Cost
retail_stocks

# Investment portfolio calculations 
port_Cost_R = sum(retail_stocks$Portfolio_Cost)
port_Value_R = sum(retail_stocks$Portfolio_Price)

# Health_position = sum(tech_stocks$Portfolio_Revenue)
Retail_position = port_Value_R - port_Cost_R
Retail_position

# Health_position - bullish 
cat("Value of Retail stocks portfolio: $", port_Value_R)
cat("Cost of Retail stocks portfolio: $", port_Cost_R)
cat("Position on Retail stocks portfolio: $", Retail_position)

# retail_plot_table - for graph 
retail_plot_table <- select(retail_stocks, symbol, date, adjusted, Portfolio_Revenue)

#graph
retail_plot <- retail_plot_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, 
                   .facet_ncol = 2, 
                   .facet_scales = "free",
                   .interactive = FALSE)
retail_plot

# add worksheet
addWorksheet(wb, sheetName = "Stock_Anysis_Retail", gridLines = FALSE)

# add plot
print(retail_plot)
insertPlot(wb, 3, width = 5, height = 3.5, fileType = "png", units = "in")

# save the workbook
saveWorkbook(wb, "005_excel_workbook.xlsx", overwrite = TRUE)
