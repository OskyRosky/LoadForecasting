inicio <- Sys.time()


##########################################################################################
##########################################################################################
#                                        US  2121                                        #
##########################################################################################
##########################################################################################

#------------------------------------------------------
#######################################################
#                 Libraries                           #
#######################################################
#------------------------------------------------------

suppressMessages(library(xlsx))
suppressMessages(library(readxl))
suppressMessages(library(lubridate))
suppressMessages(library(fpp2))
suppressMessages(library(readxl))
suppressMessages(library(highcharter))
suppressMessages(library(dplyr))
suppressMessages(library(tidyr))
suppressMessages(library(psych))
suppressMessages(library(highcharter))
suppressMessages(library(formattable))
suppressMessages(library(highcharter))
suppressMessages(library(viridisLite))
suppressMessages(library(stringi))
suppressMessages(library(data.table))
suppressMessages(library(tidyr))
suppressMessages(library(forecast))
suppressMessages(library(RPostgreSQL))
suppressMessages(library(DBI))
suppressMessages(library(dplyr))


#------------------------------------------------------
#######################################################
#                   Data                              #
#######################################################
#------------------------------------------------------

##############################
#  1. Set working directory   #  
##############################

setwd("C:/Users/Oscar Centeno/OneDrive - Enchanted Rock/US 2121/Report PBI")

########################
# 2. Import data       #
########################

df_15min <- read_excel("df_15min.xlsx", sheet = "General") 

########################################
#  3.  Change col type                 #
########################################


df_15min <- df_15min %>%  mutate(
  # Load = tsclean(Load),               ---> this time the results are not great
  Date2 = as.character.Date(Date),
  # Date = as.Date(Date, format = "%Y-%m-%d %H:%M:%S")   # "%Y-%m-%d %H:%M:%S"
)# %>% select(Date,Date2,Load)

#####################
#  4. Df_filter     #
#####################

df_15min <- df_15min %>% filter(Date >= ymd_hms("2022-08-01 00:00:00"))

# str(df_15min)
# Method 2: Using `slice()`

######################################
#  5. Imputation                     #
######################################


mean_load <- mean(df_15min$Load, na.rm = TRUE)
df_15min$Load[is.na(df_15min$Load)] <- mean_load
mean_load <- mean(df_15min$Load, na.rm = TRUE)

# sum(is.na(df_15min$Load))     ---> 0

LL <- 0.05

df_15min <- df_15min  %>% dplyr::mutate(
  Load =    ifelse(Load < LL, mean_load, df_15min$Load)
) 


######################################
#   6. Load plot                     #
######################################


#Load_plot_15min_imput <- highchart() %>% 

#  hc_title(text = "",
#           margin = 20, align = "center",
#           style = list(color = "#129", useHTML = TRUE)) %>% 
#  hc_subtitle(text = "",
#              align = "right",
#              style = list(color = "#634", fontWeight = "bold")) %>%
#  hc_credits(enabled = TRUE, # add credits
#             text = "" # href = "www.cgr.go.cr"
#  ) %>%
#  hc_legend(align = "left", verticalAlign = "top",
#            layout = "vertical", x = 0, y = 100) %>%
#  hc_exporting(enabled = TRUE) %>% 

#  hc_xAxis(categories = df_15min$Date2) %>% 
#  hc_add_series(name = "Load", data = df_15min$Load, color="blue") %>% 

#  hc_yAxis(title = list(text = "Load"), labels = list(format = "{value}"))%>% 
#  hc_xAxis(title = list(text = "Date") )  %>% 
#  hc_chart(zoomType = "xy")

#Load_plot_15min_imput


#------------------------------------------------------
#######################################################
#                   TS estimation                     #
#######################################################
#------------------------------------------------------

#################
#   1. nna      #
#################

series_15min <- ts(df_15min$Load, frequency = 4*24,  start=c(2021,257))
mod <- nnetar(series_15min) 

fitted_nna_train <- as.data.frame(as.numeric(fitted(mod))) %>%  dplyr::rename(
  `Load Estimated` = `as.numeric(fitted(mod))`
)
df_15min <- dplyr::bind_cols(df_15min,fitted_nna_train)


####################################
#   2. Observed vs Prediccted      #
####################################

# Load_plot_obs_pred <- highchart() %>% 

#  hc_title(text = "",
#           margin = 20, align = "center",
#           style = list(color = "#129", useHTML = TRUE)) %>% 
#  hc_subtitle(text = "",
#              align = "right",
#              style = list(color = "#634", fontWeight = "bold")) %>%
#  hc_credits(enabled = TRUE, # add credits
#             text = "" # href = "www.cgr.go.cr"
#  ) %>%
#  hc_legend(align = "left", verticalAlign = "top",
#            layout = "vertical", x = 0, y = 100) %>%
#  hc_exporting(enabled = TRUE) %>% 

#  hc_xAxis(categories = train_set_15min$Date2) %>% 
#  hc_add_series(name = "Load", data = train_set_15min$Load, color="blue") %>% 
#  hc_add_series(name = "Fitted NNA", data = train_set_15min$`Load Estimated`, color="green") %>% 


#  hc_yAxis(title = list(text = "Load"), labels = list(format = "{value}"))%>% 
#  hc_xAxis(title = list(text = "Date") )  %>% 
#  hc_chart(zoomType = "xy")


# Load_plot_obs_pred

#------------------------------------------------------
#######################################################
#                   Forecasting                     #
#######################################################
#------------------------------------------------------

#########################
# 1. Forecast H periods #
#########################

T <- 24
H <- 12


fcast <- forecast(mod, PI=TRUE, h=H, level = 95)
fcast <- as.data.frame(fcast)

#####################################
# 2. Dataframe with forecast values #
#####################################

# Creation of Date



Date_last<- df_15min %>% slice_tail(n = 1) %>% pull(Date) %>% ymd_hms() %>% + minutes(15)
# Date_last <- Date_last %>% ymd_hms() # %>% add_minutes(15)
# Date_last + 15
  

Date_last <- format(Date_last, "%Y-%m-%d %H:%M:%S")

start_date <- ymd_hms(Date_last)

time_seq <- seq(start_date, by = "15 min", length.out = H)



Forecast_15min <- data.frame(Date = time_seq, 
                       Forecast = as.numeric(fcast$`Point Forecast`),
                       Lo = as.numeric(fcast$`Lo 95`),
                       Hi = as.numeric(fcast$`Hi 95`)
)


# Forecasting_test$Date = as.character(Forecasting_test$Date)  

################################################################
# 3. Plotting: Observed validation                             #
################################################################


nna_val_forecast <- highchart() %>% 
  
  hc_title(text = "",
           margin = 20, align = "center",
           style = list(color = "#129", useHTML = TRUE)) %>% 
  hc_subtitle(text = "",
              align = "right",
              style = list(color = "#634", fontWeight = "bold")) %>%
  hc_credits(enabled = TRUE, # add credits
             text = "" # href = "www.cgr.go.cr"
  ) %>%
  hc_legend(align = "left", verticalAlign = "top",
            layout = "vertical", x = 0, y = 100) %>%
  hc_exporting(enabled = TRUE) %>% 
  
  
  hc_xAxis(categories = Forecast_15min$Date) %>% 
  hc_add_series(name = "Forecast", data = Forecast_15min$Forecast, color="green") %>% 
  hc_add_series(name = "IC Lo 95%", data = Forecast_15min$Lo, color="red") %>% 
  hc_add_series(name = "IC Hi 95%", data = Forecast_15min$Hi, color="red") %>% 
  #  hc_add_series(name = "Load tslm", data = test_set_forecast_hour$Forecast_tslm, color="grey") %>% 
  
  
  hc_yAxis(title = list(text = "Load"), labels = list(format = "{value}"))%>% 
  hc_xAxis(title = list(text = "Date") )  %>% 
  hc_chart(zoomType = "xy")


nna_val_forecast






################################################################
# 4. Export Data                                               #
################################################################



setwd("C:/Users/Oscar Centeno/OneDrive - Enchanted Rock/US 2121/Report PBI/Tables")  

Forecast_H = paste0("Forecasting_15min_H.xlsx")  

write.xlsx(Forecast_15min, file = Forecast_H)



################################################################
#  Tiempo de ejecución                                         #
################################################################


fin <- Sys.time()  # Capturar el tiempo de finalización
duracion <- fin - inicio 
duracion