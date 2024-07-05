install.packages("readxl")
library(readxl)
install.packages("dplyr")
library(dplyr)
install.packages("tidyr")
library(tidyr)
install.packages("car")
library(car)
install.packages("plm")
library(plm)
install.packages("glmnet")
library(glmnet)
library(lmtest)
install.packages("ggplot2")
library(ggplot2)

#-------------------------------------------------------------------------------
# ĐỌC FILE EXCEL
file_path <- "C:/Users/ASUS/Downloads/GPM2/FINAL/DataGPM2.xlsx"
data <- read_excel(file_path, sheet = 'FIELD')

#-------------------------------------------------------------------------------
# CHUYỂN ĐỔI KIỂU DỮ LIỆU
data$`Cash and Equivalents` <- as.numeric(as.character(data$`Cash and Equivalents`), 
                                          na.strings = c("", "NA", "NaN"))
data$`Total Assets, Reported` <- as.numeric(as.character(data$`Total Assets, Reported`), 
                                            na.strings = c("", "NA", "NaN"))
data$`Total Liabilities` <- as.numeric(as.character(data$`Total Liabilities`), 
                                       na.strings = c("", "NA", "NaN"))
data$`Total Equity` <- as.numeric(as.character(data$`Total Equity`), 
                                  na.strings = c("", "NA", "NaN"))
data$`Total Current Assets` <- as.numeric(as.character(data$`Total Current Assets`), 
                                          na.strings = c("", "NA", "NaN"))
data$`Total Current Liabilities` <- as.numeric(as.character(data$`Total Current Liabilities`), 
                                               na.strings = c("", "NA", "NaN"))
data$`ROE Total Equity %` <- as.numeric(as.character(data$`ROE Total Equity %`), 
                                        na.strings = c("", "NA", "NaN"))
data$`Net Income After Taxes` <- as.numeric(as.character(data$`Net Income After Taxes`), 
                                            na.strings = c("", "NA", "NaN"))
data$`Total Inventory` <- as.numeric(as.character(data$`Total Inventory`), 
                                                 na.strings = c("", "NA", "NaN"))
#-------------------------------------------------------------------------------
# TÍNH TOÁN CÁC CHỈ SỐ
data$FF <- with(data, `Cash and Equivalents` / `Total Assets, Reported` + 
                  (1 - `Total Liabilities` / `Total Assets, Reported`))
data$DE <- with(data, `Total Liabilities` / `Total Equity`)
data$QR <- with(data, (`Total Current Assets` - `Total Inventory`) / `Total Current Liabilities`)
data$SIZE <- with(data, log(`Total Assets, Reported`))
data$ROA <- with(data, `Net Income After Taxes`/ `Total Assets, Reported`)
data$INF <- data$Inflation

#-------------------------------------------------------------------------------
# THAY THẾ CÁC GIÁ TRỊ NAN BẰNG GIÁ TRỊ TRUNG BÌNH TƯƠNG ỨNG VỚI TỪNG MÃ CP
data$FF[is.na(data$FF)] <- ave(data$FF, data$Name, 
                               FUN = function(x) mean(x, na.rm = TRUE))
data$DE[is.na(data$DE)] <- ave(data$DE, data$Name, 
                                   FUN = function(x) mean(x, na.rm = TRUE))
data$QR[is.na(data$QR)] <- ave(data$QR, data$Name, 
                                           FUN = function(x) mean(x, na.rm = TRUE))
data$SIZE[is.na(data$SIZE)] <- ave(data$SIZE, data$Name, 
                                             FUN = function(x) mean(x, na.rm = TRUE))
data$ROA[is.na(data$ROA)] <- ave(data$ROA, data$Name, 
                                                 FUN = function(x) mean(x, na.rm = TRUE))
data$INF[is.na(data$INF)] <- ave(data$INF, data$Name, 
                                 FUN = function(x) mean(x, na.rm = TRUE))

#-------------------------------------------------------------------------------
# TẠO DATAFRAME MỚI VỚI CÁC CHỈ SỐ ĐÃ TÍNH
df <- data %>%
  select(`Name`, `Date`, FF, DE, QR, SIZE, ROA, INF)

#MÔ TẢ ĐẶC ĐIỂM DỮ LIỆU
summary(df)

#-------------------------------------------------------------------------------
# TÍNH HỆ SỐ TƯƠNG QUAN
cor_matrix <- cor(df[, 3:8])
cor_matrix

# VẼ MA TRẬN TƯƠNG QUAN
library(corrplot)
corrplot(cor_matrix, method = "number", main = "Ma trận tương quan", 
         addCoef.col = "black", number.cex = 0.8, mar = c(1, 1, 1, 1))

#BOX PLOT CỦA BIẾN LINH HOẠT TÀI CHÍNH
ggplot(df, aes(x = "", y = FF)) + 
  geom_boxplot() + 
  labs(x = "", y = "Linh hoạt tài chính") + 
  ggtitle("Biểu đồ hộp của biến FF") + 
  theme_classic()

# BIỂU ĐỒ TẦN SUẤT CỦA BIẾN LINH HOẠT TÀI CHÍNH
ggplot(df, aes(x = FF)) +
  geom_histogram(binwidth = 0.1, fill = "skyblue", color = "black") +
  labs(x = "Linh hoạt tài chính", y = "Tần suất") +
  ggtitle("Biểu đồ tần suất của biến Linh hoạt tài chính") +
  theme_minimal()

# BIỂU ĐỒ TẦN SUẤT CỦA BIẾN ĐO LƯỜNG HIỆU QUẢ HOẠT ĐỘNG
ggplot(df, aes(x = ROA)) +
  geom_histogram(binwidth = 0.1, fill = "lavender", color = "black") +
  labs(x = "Hiệu quả quản lý rủi ro", y = "Tần suất") +
  ggtitle("Biểu đồ tần suất của biến Hiệu quả hoạt động") +
  theme_minimal()

#-------------------------------------------------------------------------------
# THỰC HIỆN HỒI QUY BẰNG MÔ HÌNH FEM
fem_model <- plm(FF ~ DE + QR + SIZE + ROA + INF, 
                 data = df, model = "random", effect = "individual")
summary(fem_model)

#KIỂM TRA ĐA CỘNG TUYẾN
vif(fem_model)

#-------------------------------------------------------------------------------
# THỰC HIỆN HỒI QUY BẰNG MÔ HÌNH FEM Ở CÁC NĂM TẠI THỜI ĐIỂM DỊCH COVID-19
library(lubridate)
df_covid <- df %>%
  filter(year(Date) %in% c(2020, 2021))

# THÊM BIẾN SỐ
df_covid <- df_covid %>%
  mutate(cancel_public_events = ifelse(year(Date) == 2020, 1, 2),
         stay_home_requirements = ifelse(year(Date) == 2020, 2, 2),
         restriction_gatherings = ifelse(year(Date) == 2020, 3, 3),
         close_public_transport = ifelse(year(Date) == 2020, 0, 2),
         restrictions_internal_movements = ifelse(year(Date) == 2020, 0, 2),
         school_closures = ifelse(year(Date) == 2020, 3, 2),
         international_travel_controls = ifelse(year(Date) == 2020, 4, 3),
         workplace_closures = ifelse(year(Date) == 2020, 2, 2))

df_covid <- df_covid %>%
  mutate(
    SINDEX = (cancel_public_events + stay_home_requirements + close_public_transport +
              restrictions_internal_movements + school_closures + international_travel_controls + 
              restriction_gatherings + workplace_closures) / 8
  )

df_covid <- df_covid %>%
  mutate(
    cancel_public_events = cancel_public_events / 2 * 100,
    stay_home_requirements = stay_home_requirements / 2 * 100,
    close_public_transport = close_public_transport / 2 * 100,
    restrictions_internal_movements = restrictions_internal_movements / 2 * 100,
    school_closures = school_closures / 3 * 100,
    international_travel_controls = international_travel_controls / 4 * 100,
    restriction_gatherings = restriction_gatherings / 3 * 100,
    workplace_closures = workplace_closures / 2 * 100
  )

df_covid <- df_covid %>%
  mutate(
    SINDEX = (cancel_public_events + stay_home_requirements + close_public_transport +
              restrictions_internal_movements + school_closures + international_travel_controls + 
              restriction_gatherings + workplace_closures) / 8
  )


df_covid <- df_covid %>%
  mutate(
    SINDEX_ln = log(SINDEX)
  )

# THỰC HIỆN HỒI QUY BẰNG MÔ HÌNH FEM
fem_model <- plm(FF ~ DE + QR + SIZE + ROA + SINDEX_ln, 
                 data = df_covid, model = "within", effect = "individual")
summary(fem_model)
