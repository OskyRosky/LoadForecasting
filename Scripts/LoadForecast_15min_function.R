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
#               Creación de la función                #
#######################################################
#------------------------------------------------------


LoadForecast <- function(T, H) {
  #------------------------------------------------------
  #######################################################
  #           Splitting  in train - validatiion         #
  #######################################################
  #------------------------------------------------------
  
  
  # dim(df_15min)[1]
  
  T <- T
  H <- H
  Diff_24 <- dim(df_15min)[1]-T
  #Diff_24
  
  #split_index <- floor(0.8 * nrow(df_hours))
  split_index_hour <- floor(Diff_24)
  
  
  train_set_15min <- df_15min[1:split_index_hour, ]
  test_set_15min <- df_15min[(split_index_hour + 1):(split_index_hour + H), ]
  
  
  # Origin
  # train_set_15min <- df_15min[1:split_index_hour, ]
  # test_set_15min <- df_15min[(split_index_hour + 1):nrow(df_15min), ]
  
  #paste0("Len hour_train_set"," ", dim(min15_train_set)[1],"---", "Len hour_test_set", " ",dim(min15_test_set)[1] )
  
  #------------------------------------------------------
  #######################################################
  #                   TS estimation                     #
  #######################################################
  #------------------------------------------------------
  
  #################
  #   1. nna      #
  #################
  
  series_15min <- ts(train_set_15min$Load, frequency = 4*24,  start=c(2021,257))
  mod <- nnetar(series_15min) 
  
  fitted_nna_train <- as.data.frame(as.numeric(fitted(mod))) %>%  dplyr::rename(
    `Load Estimated` = `as.numeric(fitted(mod))`
  )
  train_set_15min <- dplyr::bind_cols(train_set_15min,fitted_nna_train)
  
  
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
  
  ####################################
  #   3. Godness of fit              #
  ####################################
  
  
  #########
  #  MAE  #
  #########
  
  mae <- function(real, predicho) {
    # Eliminar las filas donde 'real' o 'predicho' son NA
    indices_validos <- !(is.na(real) | is.na(predicho))
    real <- real[indices_validos]
    predicho <- predicho[indices_validos]
    
    # Calcular el MAE
    mean(abs(real - predicho), na.rm = TRUE)
  }
  
  mae <-  mae(real = train_set_15min$Load, predicho = train_set_15min$`Load Estimated`)
  
  #########
  #  MSE  #
  #########
  
  # Función para calcular el Mean Squared Error (MSE)
  mse <- function(real, predicho) {
    # Eliminar las filas donde 'real' o 'predicho' son NA
    indices_validos <- !(is.na(real) | is.na(predicho))
    real <- real[indices_validos]
    predicho <- predicho[indices_validos]
    
    # Calcular el MSE
    mean((real - predicho)^2, na.rm = TRUE)
  }
  
  mse <-  mse(real = train_set_15min$Load, predicho = train_set_15min$`Load Estimated`)
  
  
  ##########
  #  RMSe  #
  ##########
  
  # Función para calcular el Root Mean Squared Error (RMSE)
  rmse <- function(real, predicho) {
    # Eliminar las filas donde 'real' o 'predicho' son NA
    indices_validos <- !(is.na(real) | is.na(predicho))
    real <- real[indices_validos]
    predicho <- predicho[indices_validos]
    
    # Calcular el RMSE
    sqrt(mean((real - predicho)^2, na.rm = TRUE))
  }
  
  # Uso de la función
  rmse <-  rmse(real = train_set_15min$Load, predicho = train_set_15min$`Load Estimated`)
  
  ##########
  #  MAPE  #
  ##########
  
  # Reemplazamos los NA con la media en los datos reales
  prom_load <- mean(train_set_15min$Load, na.rm = TRUE)
  
  # Calculamos el MAPE
  mape <- mean(abs((prom_load - train_set_15min$`Load Estimated`) / train_set_15min$`Load Estimated`), na.rm = TRUE) * 100
  
  
  
  
  ###################################
  #  Table of godness of fit        #
  ################################### 
  
  GOF <- data.frame ( MAE = mae,
                      MSE = mse,
                      RMSE = rmse,
                      MAPE = mape
  )
  
  #------------------------------------------------------
  #######################################################
  #                   Forecasting                     #
  #######################################################
  #------------------------------------------------------
  
  #########################
  # 1. Forecast H periods #
  #########################
  
  fcast <- forecast(mod, PI=TRUE, h=H, level = 95)
  fcast <- as.data.frame(fcast)
  
  #####################################
  # 2. Dataframe with forecast values #
  #####################################
  
  # Creation of Date
  
  Date_first_valset <- as.data.frame(test_set_15min[1,1]) 
  Date_first_valset <- format(Date_first_valset$Date, "%Y-%m-%d %H:%M:%S")
  
  start_date <- ymd_hms(Date_first_valset)
  
  time_seq <- seq(start_date, by = "15 min", length.out = H)
  
  
  
  Forecast <- data.frame(Date = time_seq, 
                         Forecast = as.numeric(fcast$`Point Forecast`),
                         Lo = as.numeric(fcast$`Lo 95`),
                         Hi = as.numeric(fcast$`Hi 95`)
  )
  
  #########################################################################
  # 3. Merge Forecasting with Loading validation set with forecast values #
  #########################################################################
  
  
  Forecasting <- merge(test_set_15min, Forecast, by = "Date") %>% 
    dplyr::select(Date, Date2, Load, Forecast, Lo, Hi)
  
  ################################################################
  # 4. Plotting: Observed validation                             #
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
    
    
    hc_xAxis(categories = Forecasting$Date2) %>% 
    hc_add_series(name = "Load", data = Forecasting$Load, color="blue") %>% 
    hc_add_series(name = "Forecast", data = Forecasting$Forecast, color="green") %>% 
    hc_add_series(name = "IC Lo 95%", data = Forecasting$Lo, color="red") %>% 
    hc_add_series(name = "IC Hi 95%", data = Forecasting$Hi, color="red") %>% 
    #  hc_add_series(name = "Load tslm", data = test_set_forecast_hour$Forecast_tslm, color="grey") %>% 
    
    
    hc_yAxis(title = list(text = "Load"), labels = list(format = "{value}"))%>% 
    hc_xAxis(title = list(text = "Date") )  %>% 
    hc_chart(zoomType = "xy")
  
  
  nna_val_forecast

  #############################
  #    Export Forecasting     #
  #############################  
  
  
setwd("C:/Users/Oscar Centeno/OneDrive - Enchanted Rock/US 2121/Report PBI/Tables")  
  
Forecast_accuracy = paste0("Forecasting_15min_accuracy.xlsx")  
Godness_of_fit = paste0("Godness_of_fit_15min.xlsx")
  
write.xlsx(Forecasting, file = Forecast_accuracy)
write.xlsx(GOF, file = Godness_of_fit)  
  
  
#return(nna_val_forecast)
  
return(list( Forecasting, nna_val_forecast, GOF))
  

  


}


LoadForecast(T=36,H=36)



fin <- Sys.time()  # Capturar el tiempo de finalización
duracion <- fin - inicio 
duracion
