# Load the necessary libraries
library(readxl)
library(writexl)
library(dplyr)
library(reshape2)
library(xlsx)
library(ggplot2)
library(plotly)

#----------------------------과제데이터(정부연구비)---------------------------------

# List of Excel file names (change these to your file names)
file_names <- c("NTIS_data/Projects_2007.xls", "NTIS_data/Projects_2008.xls", "NTIS_data/Projects_2009.xls",
                "NTIS_data/Projects_2010.xls", "NTIS_data/Projects_2011.xls", "NTIS_data/Projects_2012.xls",
                "NTIS_data/Projects_2013.xls", "NTIS_data/Projects_2014.xls", "NTIS_data/Projects_2015.xls",
                "NTIS_data/Projects_2016.xls", "NTIS_data/Projects_2017.xlsx","NTIS_data/Projects_2018.xlsx",
                "NTIS_data/Projects_2019.xlsx","NTIS_data/Projects_2020.xlsx","NTIS_data/Projects_2021.xlsx")

# Specify the columns you want to extract (change these to your column names)
columns_to_extract <- c("과제수행년도", "과제고유번호", "연구개발단계코드", "과학기술표준분류코드1-대",
                        "과학기술표준분류가중치1", "과학기술표준분류코드2-대", "과학기술표준분류가중치2",
                        "과학기술표준분류코드3-대", "과학기술표준분류가중치3", "정부연구비(원)")

# Initialize an empty data frame to store the merged data
merged_data <- data.frame()

# 2013년 개정된 과학기술표준분류를 2007~2012년 데이터에 반영하기 위한 변환표 로딩
find_replace_df <- read_excel("표준분류변환.xlsx", col_types = "text")

# Loop through each Excel file and extract the specified columns
for (file in file_names) {
  # Read the Excel sheet into a data frame
  sheet_data <- read_excel(file, col_types = "text")
  
  # Select only the columns you want to extract
  selected_data <- sheet_data %>%
    select(all_of(columns_to_extract))
  
  # 과학기술표준분류 없는 row 삭제
  filtered_data <- selected_data %>%
    filter(!is.na(selected_data$"과학기술표준분류코드1-대"))
  
  # 컬럼명에 특수문자가 있으면 에러나므로 '-', (원)' 제거
  names(filtered_data) <- c("과제수행년도", "과제고유번호", "연구개발단계코드", "과학기술표준분류코드1대",
                            "과학기술표준분류가중치1", "과학기술표준분류코드2대", "과학기술표준분류가중치2",
                            "과학기술표준분류코드3대", "과학기술표준분류가중치3", "정부연구비")
  
  filtered_data <- filtered_data %>%
    mutate(
      과제수행년도 = as.numeric(과제수행년도),
      과학기술표준분류가중치1 = as.numeric(과학기술표준분류가중치1),
      과학기술표준분류가중치2 = as.numeric(과학기술표준분류가중치2),
      과학기술표준분류가중치3 = as.numeric(과학기술표준분류가중치3),
      정부연구비 = as.numeric(정부연구비)
    )
  
  # 2012년 이전 데이터는 과학기술표준분류 신분류로 변경
  if (filtered_data$과제수행년도[1] <= 2012) {
    for (i in 1:nrow(find_replace_df)) {
      filtered_data["과학기술표준분류코드1대"][filtered_data["과학기술표준분류코드1대"] == as.character(find_replace_df[i, 1])] <- as.character(find_replace_df[i, 2])
      filtered_data["과학기술표준분류코드2대"][filtered_data["과학기술표준분류코드2대"] == as.character(find_replace_df[i, 1])] <- as.character(find_replace_df[i, 2])
      filtered_data["과학기술표준분류코드3대"][filtered_data["과학기술표준분류코드3대"] == as.character(find_replace_df[i, 1])] <- as.character(find_replace_df[i, 2])
    }
  }
  
  # Bind the selected data to the merged_data data frame
  if (nrow(merged_data) == 0) {
    merged_data <- filtered_data
  } else {
    merged_data <- bind_rows(merged_data, filtered_data)
  }
}

projects_data_raw <- merged_data

save(projects_data_raw, file = "Projects_merge_raw.rda")

# 과학기술표준분류 가중치 계산해서 피봇테이블 생성할 수 있게 함
temp2 <- projects_data_raw %>%
  filter(!is.na(projects_data_raw$"과학기술표준분류코드2대"))

temp3 <- projects_data_raw %>%
  filter(!is.na(projects_data_raw$"과학기술표준분류코드3대"))

temp2 <- temp2 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드2대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치2
  )

temp3 <- temp3 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드3대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치3
  )

project_data <- bind_rows(projects_data_raw, temp2)

project_data <- bind_rows(projects_data_raw, temp3)

project_data <- project_data %>%
  mutate(
    가중치적용연구비 = 정부연구비 * 과학기술표준분류가중치1 / 100
  )

# 정부연구비 NA 값인 것 제외
project_data <- project_data %>%
  filter(!is.na(project_data$"정부연구비"))

# 최종결과 저장
write_xlsx(project_data, "Projects_merge.xlsx", col_names=TRUE)
save(project_data, file = "Projects_merge.rda")

# 연도별로, 표준분류-연구개발단계 피봇테이블 생성
for (i in 2007:2021) {
  pivot_table <- project_data %>% 
    filter(project_data["과제수행년도"] == i) %>% 
    dcast(과학기술표준분류코드1대 ~ 연구개발단계코드, sum , value.var = "정부연구비")
  
  write.xlsx(pivot_table, "Projects_pivot.xlsx", sheetName = as.character(i), append = TRUE, row.names = FALSE)
}


#--------------------------성과데이터(사업화 건수)----------------------------------------

# List of Excel file names (change these to your file names)
file_names <- c("NTIS_data/Com_2007.xlsx","NTIS_data/Com_2008.xlsx","NTIS_data/Com_2009.xlsx",
                "NTIS_data/Com_2010.xlsx","NTIS_data/Com_2011.xlsx","NTIS_data/Com_2012.xls",
                "NTIS_data/Com_2013.xls", "NTIS_data/Com_2014.xls", "NTIS_data/Com_2015.xlsx",
                "NTIS_data/Com_2016.xlsx","NTIS_data/Com_2017.xlsx","NTIS_data/Com_2018.xlsx",
                "NTIS_data/Com_2019.xlsx","NTIS_data/Com_2020.xlsx","NTIS_data/Com_2021.xlsx")

# Specify the columns you want to extract (change these to your column names)
columns_to_extract <- c("과제수행년도", "성과발생년도", "과제고유번호")

# Initialize an empty data frame to store the merged data
merged_data <- data.frame()

# Loop through each Excel file and extract the specified columns
for (file in file_names) {
  # Read the Excel sheet into a data frame
  sheet_data <- read_excel(file, col_types = "text")
  
  # Select only the columns you want to extract
  selected_data <- sheet_data %>%
    select(all_of(columns_to_extract))
  
  if (selected_data$과제수행년도[1] == "aa010")
    selected_data <- selected_data[-1,]
  
  selected_data <- selected_data %>%
    mutate(
      과제수행년도 = as.numeric(과제수행년도),
      성과발생년도 = as.numeric(성과발생년도)
    )
  
  # Bind the selected data to the merged_data data frame
  if (nrow(merged_data) == 0) {
    merged_data <- selected_data
  } else {
    merged_data <- bind_rows(merged_data, selected_data)
  }
}

com_data <- left_join(merged_data, projects_data_raw, "과제고유번호")

save(com_data, file = "Com_merge_raw.rda")

# 과학기술표준분류 가중치 계산해서 피봇테이블 생성할 수 있게 함
temp2 <- com_data %>%
  filter(!is.na(com_data$"과학기술표준분류코드2대"))

temp3 <- com_data %>%
  filter(!is.na(com_data$"과학기술표준분류코드3대"))

temp2 <- temp2 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드2대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치2
  )

temp3 <- temp3 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드3대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치3
  )

com_data <- bind_rows(com_data, temp2)
com_data <- bind_rows(com_data, temp3)

com_data <- com_data %>%
  mutate(
    가중치적용사업화 = 1 * 과학기술표준분류가중치1 / 100
  )

# 성과 NA 값인 것 제외
com_data <- com_data %>%
  filter(!is.na(com_data$"가중치적용사업화"))

# 최종결과 저장
write_xlsx(com_data, "Com_merge.xlsx", col_names=TRUE)
save(com_data, file = "Com_merge.rda")

# 연도별로, 표준분류-연구개발단계 피봇테이블 생성
for (i in 2007:2021) {
  pivot_table <- com_data %>% 
    filter(com_data["성과발생년도"] == i) %>% 
    dcast(과학기술표준분류코드1대 ~ 연구개발단계코드, sum , value.var = "가중치적용사업화")
  
  write.xlsx(pivot_table, "Com_pivot.xlsx", sheetName = as.character(i), append = TRUE, row.names = FALSE)
}


#--------------------------성과데이터(국내특허 건수)----------------------------------------

# List of Excel file names (change these to your file names)
file_names <- c("NTIS_data/DomPat_2007.xlsx","NTIS_data/DomPat_2008.xlsx","NTIS_data/DomPat_2009.xlsx",
                "NTIS_data/DomPat_2010.xlsx","NTIS_data/DomPat_2011.xlsx","NTIS_data/DomPat_2012.xlsx",
                "NTIS_data/DomPat_2013.xlsx","NTIS_data/DomPat_2014.xlsx","NTIS_data/DomPat_2015.xlsx",
                "NTIS_data/DomPat_2016.xlsx","NTIS_data/DomPat_2017.xlsx","NTIS_data/DomPat_2018.xlsx",
                "NTIS_data/DomPat_2019.xlsx","NTIS_data/DomPat_2020.xlsx","NTIS_data/DomPat_2021.xlsx")

# Specify the columns you want to extract (change these to your column names)
columns_to_extract <- c("과제수행년도", "성과발생년도", "과제고유번호", "특허_기여율_확정")

# Initialize an empty data frame to store the merged data
merged_data <- data.frame()

# Loop through each Excel file and extract the specified columns
for (file in file_names) {
  # Read the Excel sheet into a data frame
  sheet_data <- read_excel(file, col_types = "text")
  
  # Select only the columns you want to extract
  selected_data <- sheet_data %>%
    select(all_of(columns_to_extract))
  
  if (selected_data$과제수행년도[1] == "aa010")
    selected_data <- selected_data[-1,]
  
  selected_data <- selected_data %>%
    mutate(
      과제수행년도 = as.numeric(과제수행년도),
      성과발생년도 = as.numeric(성과발생년도),
      특허_기여율_확정 = as.numeric(특허_기여율_확정)
    )
  
  # Bind the selected data to the merged_data data frame
  if (nrow(merged_data) == 0) {
    merged_data <- selected_data
  } else {
    merged_data <- bind_rows(merged_data, selected_data)
  }
}

dompat_data <- left_join(merged_data, projects_data_raw, "과제고유번호")

save(dompat_data, file = "DomPat_merge_raw.rda")

# 과학기술표준분류 가중치 계산해서 피봇테이블 생성할 수 있게 함
temp2 <- dompat_data %>%
  filter(!is.na(dompat_data$"과학기술표준분류코드2대"))

temp3 <- dompat_data %>%
  filter(!is.na(dompat_data$"과학기술표준분류코드3대"))

temp2 <- temp2 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드2대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치2
  )

temp3 <- temp3 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드3대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치3
  )

dompat_data <- bind_rows(dompat_data, temp2)
dompat_data <- bind_rows(dompat_data, temp3)

dompat_data <- dompat_data %>%
  mutate(
    가중치적용국내특허 = 1 * 과학기술표준분류가중치1 / 100 * 특허_기여율_확정 / 100
  )

# 성과 NA 값인 것 제외
dompat_data <- dompat_data %>%
  filter(!is.na(dompat_data$"가중치적용국내특허"))

# 최종결과 저장
write_xlsx(dompat_data, "DomPat_merge.xlsx", col_names=TRUE)
save(dompat_data, file = "DomPat_merge.rda")

# 연도별로, 표준분류-연구개발단계 피봇테이블 생성
for (i in 2007:2021) {
  pivot_table <- dompat_data %>% 
    filter(dompat_data["성과발생년도"] == i) %>% 
    dcast(과학기술표준분류코드1대 ~ 연구개발단계코드, sum , value.var = "가중치적용국내특허")
  
  write.xlsx(pivot_table, "DomPat_pivot.xlsx", sheetName = as.character(i), append = TRUE, row.names = FALSE)
}


#--------------------------성과데이터(국외특허 건수)----------------------------------------

# List of Excel file names (change these to your file names)
file_names <- c("NTIS_data/IntPat_2007.xlsx","NTIS_data/IntPat_2008.xlsx","NTIS_data/IntPat_2009.xlsx",
                "NTIS_data/IntPat_2010.xlsx","NTIS_data/IntPat_2011.xlsx","NTIS_data/IntPat_2012.xls",
                "NTIS_data/IntPat_2013.xls", "NTIS_data/IntPat_2014.xls" ,"NTIS_data/IntPat_2015.xlsx",
                "NTIS_data/IntPat_2016.xls", "NTIS_data/IntPat_2017.xlsx","NTIS_data/IntPat_2018.xlsx",
                "NTIS_data/IntPat_2019.xlsx","NTIS_data/IntPat_2020.xlsx","NTIS_data/IntPat_2021.xlsx")

# Specify the columns you want to extract (change these to your column names)
columns_to_extract <- c("과제수행년도", "성과발생년도", "과제고유번호", "특허_기여율_확정")

# Initialize an empty data frame to store the merged data
merged_data <- data.frame()

# Loop through each Excel file and extract the specified columns
for (file in file_names) {
  # Read the Excel sheet into a data frame
  sheet_data <- read_excel(file, col_types = "text")
  
  # Select only the columns you want to extract
  selected_data <- sheet_data %>%
    select(all_of(columns_to_extract))
  
  if (selected_data$과제수행년도[1] == "aa010")
    selected_data <- selected_data[-1,]
  
  selected_data <- selected_data %>%
    mutate(
      과제수행년도 = as.numeric(과제수행년도),
      성과발생년도 = as.numeric(성과발생년도),
      특허_기여율_확정 = as.numeric(특허_기여율_확정)
    )
  
  # Bind the selected data to the merged_data data frame
  if (nrow(merged_data) == 0) {
    merged_data <- selected_data
  } else {
    merged_data <- bind_rows(merged_data, selected_data)
  }
}

intpat_data <- left_join(merged_data, projects_data_raw, "과제고유번호")

save(intpat_data, file = "IntPat_merge_raw.rda")

# 과학기술표준분류 가중치 계산해서 피봇테이블 생성할 수 있게 함
temp2 <- intpat_data %>%
  filter(!is.na(intpat_data$"과학기술표준분류코드2대"))

temp3 <- intpat_data %>%
  filter(!is.na(intpat_data$"과학기술표준분류코드3대"))

temp2 <- temp2 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드2대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치2
  )

temp3 <- temp3 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드3대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치3
  )

intpat_data <- bind_rows(intpat_data, temp2)
intpat_data <- bind_rows(intpat_data, temp3)

intpat_data <- intpat_data %>%
  mutate(
    가중치적용국외특허 = 1 * 과학기술표준분류가중치1 / 100 * 특허_기여율_확정 / 100
  )

# 성과 NA 값인 것 제외
intpat_data <- intpat_data %>%
  filter(!is.na(intpat_data$"가중치적용국외특허"))

# 최종결과 저장
write_xlsx(intpat_data, "IntPat_merge.xlsx", col_names=TRUE)
save(intpat_data, file = "IntPat_merge.rda")

# 연도별로, 표준분류-연구개발단계 피봇테이블 생성
for (i in 2007:2021) {
  pivot_table <- intpat_data %>% 
    filter(intpat_data["성과발생년도"] == i) %>% 
    dcast(과학기술표준분류코드1대 ~ 연구개발단계코드, sum , value.var = "가중치적용국외특허")

  write.xlsx(pivot_table, "IntPat_pivot.xlsx", sheetName = as.character(i), append = TRUE, row.names = FALSE)
}



#--------------------------성과데이터(기술이전 기술료(원))----------------------------------------

# List of Excel file names (change these to your file names)
file_names <- c("NTIS_data/License_2007.xlsx","NTIS_data/License_2008.xlsx","NTIS_data/License_2009.xlsx",
                "NTIS_data/License_2010.xlsx","NTIS_data/License_2011.xlsx","NTIS_data/License_2012.xls",
                "NTIS_data/License_2013.xls", "NTIS_data/License_2014.xls" ,"NTIS_data/License_2015.xls",
                "NTIS_data/License_2016.xlsx","NTIS_data/License_2017.xlsx","NTIS_data/License_2018.xlsx",
                "NTIS_data/License_2019.xlsx","NTIS_data/License_2020.xlsx","NTIS_data/License_2021.xlsx")

# Specify the columns you want to extract (change these to your column names)
columns_to_extract <- c("과제수행년도", "성과발생년도", "과제고유번호", "기술료_당해연도기술료_원")

# Initialize an empty data frame to store the merged data
merged_data <- data.frame()

# Loop through each Excel file and extract the specified columns
for (file in file_names) {
  # Read the Excel sheet into a data frame
  sheet_data <- read_excel(file, col_types = "text")
  
  # Select only the columns you want to extract
  selected_data <- sheet_data %>%
    select(all_of(columns_to_extract))
  
  if (selected_data$과제수행년도[1] == "aa010")
    selected_data <- selected_data[-1,]
  
  selected_data <- selected_data %>%
    mutate(
      과제수행년도 = as.numeric(과제수행년도),
      성과발생년도 = as.numeric(성과발생년도),
      기술료_당해연도기술료_원 = as.numeric(기술료_당해연도기술료_원)
    )
  
  # Bind the selected data to the merged_data data frame
  if (nrow(merged_data) == 0) {
    merged_data <- selected_data
  } else {
    merged_data <- bind_rows(merged_data, selected_data)
  }
}

load("Projects_merge.rda")

license_data <- left_join(merged_data, projects_data_raw, "과제고유번호")

save(license_data, file = "License_merge_raw.rda")

# 과학기술표준분류 가중치 계산해서 피봇테이블 생성할 수 있게 함
temp2 <- license_data %>%
  filter(!is.na(license_data$"과학기술표준분류코드2대"))

temp3 <- license_data %>%
  filter(!is.na(license_data$"과학기술표준분류코드3대"))

temp2 <- temp2 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드2대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치2
  )

temp3 <- temp3 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드3대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치3
  )

license_data <- bind_rows(license_data, temp2)

license_data <- bind_rows(license_data, temp3)

license_data <- license_data %>%
  mutate(
    가중치적용기술료 = 기술료_당해연도기술료_원 * 과학기술표준분류가중치1 / 100
  )

# 성과 NA 값인 것 제외
license_data <- license_data %>%
  filter(!is.na(license_data$"가중치적용기술료"))

# 최종결과 저장
write_xlsx(license_data, "License_merge.xlsx", col_names=TRUE)
save(license_data, file = "License_merge.rda")

# 연도별로, 표준분류-연구개발단계 피봇테이블 생성
for (i in 2007:2021) {
  pivot_table <- license_data %>% 
    filter(license_data["성과발생년도"] == i) %>% 
    dcast(과학기술표준분류코드1대 ~ 연구개발단계코드, sum , value.var = "가중치적용기술료")
  
  write.xlsx(pivot_table, "License_pivot.xlsx", sheetName = as.character(i), append = TRUE, row.names = FALSE)
}


#--------------------------성과데이터(SCIE 논문 건수)----------------------------------------

# List of Excel file names (change these to your file names)
file_names <- c("NTIS_data/SCIE_2007.xlsx","NTIS_data/SCIE_2008.xlsx","NTIS_data/SCIE_2009.xlsx",
                "NTIS_data/SCIE_2010.xlsx","NTIS_data/SCIE_2011.xlsx","NTIS_data/SCIE_2012.xlsx",
                "NTIS_data/SCIE_2013.xlsx","NTIS_data/SCIE_2014.xlsx","NTIS_data/SCIE_2015.xlsx",
                "NTIS_data/SCIE_2016.xlsx","NTIS_data/SCIE_2017.xlsx","NTIS_data/SCIE_2018.xlsx",
                "NTIS_data/SCIE_2019.xlsx","NTIS_data/SCIE_2020.xlsx","NTIS_data/SCIE_2021.xlsx")

# Specify the columns you want to extract (change these to your column names)
columns_to_extract <- c("과제수행년도", "성과발생년도", "과제고유번호", "논문_기여율_확정")

# Initialize an empty data frame to store the merged data
merged_data <- data.frame()

# Loop through each Excel file and extract the specified columns
for (file in file_names) {
  # Read the Excel sheet into a data frame
  sheet_data <- read_excel(file, col_types = "text")
  
  # Select only the columns you want to extract
  selected_data <- sheet_data %>%
    select(all_of(columns_to_extract))
  
  if (selected_data$과제수행년도[1] == "aa010")
    selected_data <- selected_data[-1,]
  
  selected_data <- selected_data %>%
    mutate(
      과제수행년도 = as.numeric(과제수행년도),
      성과발생년도 = as.numeric(성과발생년도),
      논문_기여율_확정 = as.numeric(논문_기여율_확정)
    )
  
  # Bind the selected data to the merged_data data frame
  if (nrow(merged_data) == 0) {
    merged_data <- selected_data
  } else {
    merged_data <- bind_rows(merged_data, selected_data)
  }
}

load("Projects_merge.rda")

SCIE_data <- left_join(merged_data, projects_data_raw, "과제고유번호")

save(SCIE_data, file = "SCIE_merge_raw.rda")

# 과학기술표준분류 가중치 계산해서 피봇테이블 생성할 수 있게 함
temp2 <- SCIE_data %>%
  filter(!is.na(SCIE_data$"과학기술표준분류코드2대"))

temp3 <- SCIE_data %>%
  filter(!is.na(SCIE_data$"과학기술표준분류코드3대"))

temp2 <- temp2 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드2대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치2
  )

temp3 <- temp3 %>%
  mutate(
    과학기술표준분류코드1대 = 과학기술표준분류코드3대,
    과학기술표준분류가중치1 = 과학기술표준분류가중치3
  )

SCIE_data <- bind_rows(SCIE_data, temp2)

SCIE_data <- bind_rows(SCIE_data, temp3)

SCIE_data <- SCIE_data %>%
  mutate(
    가중치적용SCIE = 1 * 과학기술표준분류가중치1 / 100 * 논문_기여율_확정 / 100
  )

# 성과 NA 값인 것 제외
SCIE_data <- SCIE_data %>%
  filter(!is.na(SCIE_data$"가중치적용SCIE"))

# 엑셀 한계치인 100만 row 넘어서, 2015년 이전 데이터 분리
SCIE_07_15 <- SCIE_data %>%
  filter(SCIE_data["성과발생년도"] <= 2015)
SCIE_16_21 <- SCIE_data %>%
  filter(SCIE_data["성과발생년도"] >= 2016)

# 최종결과 저장
write_xlsx(SCIE_07_15, "SCIE_07-15.xlsx", col_names=TRUE)
write_xlsx(SCIE_16_21, "SCIE_16-21.xlsx", col_names=TRUE)
save(SCIE_data, file = "SCIE_merge.rda")

# 연도별로, 표준분류-연구개발단계 피봇테이블 생성
for (i in 2007:2021) {
  pivot_table <- SCIE_data %>% 
    filter(SCIE_data["성과발생년도"] == i) %>% 
    dcast(과학기술표준분류코드1대 ~ 연구개발단계코드, sum , value.var = "가중치적용SCIE")
  
  write.xlsx(pivot_table, "SCIE_pivot.xlsx", sheetName = as.character(i), append = TRUE, row.names = FALSE)
}


#--------------------------2단계 분석----------------------------------------

# melt function 을 이용해서 각 파일의 여러 시트를 하나의 테이블로 합침

filenames <- c("Projects_pivot.xlsx", "Com_pivot.xlsx", "DomPat_pivot.xlsx", "IntPat_pivot.xlsx", "License_pivot.xlsx", "SCIE_pivot.xlsx")
tablenames <- c("Projects", "Com", "DomPat", "IntPat", "License", "SCIE")

melted_output <- list()

for (j in 1:length(filenames)) {
  excel_file <- filenames[j]
  sheet_names <- excel_sheets(excel_file)
  
  melted_data <- data.frame()
  
  melted_data <- read_excel(excel_file, sheet = sheet_names[1])
  melted_data <- melt(melted_data, id.vars = c("과학기술표준분류코드1대"))
  melted_data$과학기술표준분류코드1대 <- paste(melted_data$과학기술표준분류코드1대, "-", melted_data$variable, sep = "")
  melted_data <- melted_data[,-2]
  colnames(melted_data) <- c('Pair', sheet_names[1])
  
  # Iterate through each sheet and apply the melt function
  for (i in 2:length(sheet_names)) {
    sheet_data <- read_excel(excel_file, sheet = sheet_names[i])
    # Apply the melt function to convert the data frame to long format
    melted <- melt(sheet_data, id.vars = c("과학기술표준분류코드1대"))
    melted$과학기술표준분류코드1대 <- paste(melted$과학기술표준분류코드1대, "-", melted$variable, sep = "")
    melted <- melted[,-2]
    
    colnames(melted) <- c('Pair', sheet_names[i])
    
    melted_data <- merge(melted_data, melted, by = "Pair", all = TRUE)
  }
  
  melted_data <- melted_data %>% replace(is.na(.), 0)
  melted_output[[tablenames[j]]] <- melted_data
}

outputnames <- c("Projects_melt.xlsx", "Com_melt.xlsx", "DomPat_melt.xlsx", "IntPat_melt.xlsx", "License_melt.xlsx", "SCIE_melt.xlsx")

for (j in 1:length(filenames)) {
  write.xlsx(melted_output[[j]], outputnames[j], row.names = FALSE)
}


# 그래프를 그릴 수 있는 형태의 테이블을 각 성과에 대하여 생성함
# Transpose, normalize 는 엑셀로 한 다음에 불러와서 작업

t_names <- c("Projects_melt_t_n.xlsx", "Com_melt_t_n.xlsx", "DomPat_melt_t_n.xlsx", "IntPat_melt_t_n.xlsx", "License_melt_t_n.xlsx", "SCIE_melt_t_n.xlsx")

merged <- read_excel(t_names[1]) %>% replace(is.na(.), 0) %>% melt(id.vars = c('Pair'))
merged <- cbind(paste(merged$Pair, "-", merged$variable, sep = ""), merged)
colnames(merged) <- c('Code', 'Year', 'Pair', tablenames[1])


# Time-series connected scatter plot 그리기
# https://stackoverflow.com/questions/61285157/time-series-connected-scatter-plot-in-r-as-image-attached
# https://stackoverflow.com/questions/60778893/how-to-apply-geom-label-to-the-last-value-data-point

# 사업화 건수
transposed <- read_excel(t_names[2]) %>% replace(is.na(.), 0) %>% melt(id.vars = c('Pair'))
transposed <- cbind(paste(transposed$Pair, "-", transposed$variable, sep = ""), transposed)
colnames(transposed) <- c('Code', 'Year', 'Pair', tablenames[2])
transposed <- transposed[,-2]
transposed <- transposed[,-2]

p <- merge(merged, transposed, by = "Code", all = TRUE)
plt <- p %>%
  ggplot(aes(x = Projects, y = Com, group = Pair, color = Pair, dlabel = Year)) + 
  geom_point(size = 2) +
  ggtitle("사업화 건수") + xlab("정부연구비(normalized)") + ylab("사업화 건수(normalized)") +
  geom_text(data = filter(p, Year == 2021), aes(label = Pair), size = 5, nudge_x = 0.06) +
  geom_text(data = p, aes(label = Year), hjust = 1, vjust = 0.5, nudge_y = 0.03) +
  geom_path()

ggplotly(plt)

# 국내특허 건수
transposed <- read_excel(t_names[3]) %>% replace(is.na(.), 0) %>% melt(id.vars = c('Pair'))
transposed <- cbind(paste(transposed$Pair, "-", transposed$variable, sep = ""), transposed)
colnames(transposed) <- c('Code', 'Year', 'Pair', tablenames[3])
transposed <- transposed[,-2]
transposed <- transposed[,-2]

p <- merge(merged, transposed, by = "Code", all = TRUE)
plt <- p %>%
  ggplot(aes(x = Projects, y = DomPat, group = Pair, color = Pair, dlabel = Year)) + 
  geom_point(size = 2) +
  ggtitle("국내특허 건수") + xlab("정부연구비(normalized)") + ylab("국내특허 건수(normalized)") +
  geom_text(data = filter(p, Year == 2021), aes(label = Pair), size = 5, nudge_x = 0.06) +
  geom_text(data = p, aes(label = Year), hjust = 1, vjust = 0.5, nudge_y = 0.03) +
  geom_path()

ggplotly(plt)

# 해외특허 건수
transposed <- read_excel(t_names[4]) %>% replace(is.na(.), 0) %>% melt(id.vars = c('Pair'))
transposed <- cbind(paste(transposed$Pair, "-", transposed$variable, sep = ""), transposed)
colnames(transposed) <- c('Code', 'Year', 'Pair', tablenames[4])
transposed <- transposed[,-2]
transposed <- transposed[,-2]

p <- merge(merged, transposed, by = "Code", all = TRUE)
plt <- p %>%
  ggplot(aes(x = Projects, y = IntPat, group = Pair, color = Pair, dlabel = Year)) + 
  geom_point(size = 2) +
  ggtitle("해외특허 건수") + xlab("정부연구비(normalized)") + ylab("해외특허 건수(normalized)") +
  geom_text(data = filter(p, Year == 2021), aes(label = Pair), size = 5, nudge_x = 0.06) +
  geom_text(data = p, aes(label = Year), hjust = 1, vjust = 0.5, nudge_y = 0.03) +
  geom_path()

ggplotly(plt)

# 기술이전 액수
transposed <- read_excel(t_names[5]) %>% replace(is.na(.), 0) %>% melt(id.vars = c('Pair'))
transposed <- cbind(paste(transposed$Pair, "-", transposed$variable, sep = ""), transposed)
colnames(transposed) <- c('Code', 'Year', 'Pair', tablenames[5])
transposed <- transposed[,-2]
transposed <- transposed[,-2]

p <- merge(merged, transposed, by = "Code", all = TRUE)
plt <- p %>%
  ggplot(aes(x = Projects, y = License, group = Pair, color = Pair, dlabel = Year)) + 
  geom_point(size = 2) +
  ggtitle("기술이전 기술료") + xlab("정부연구비(normalized)") + ylab("기술이전 기술료(normalized)") +
  geom_text(data = filter(p, Year == 2021), aes(label = Pair), size = 5, nudge_x = 0.06) +
  geom_text(data = p, aes(label = Year), hjust = 1, vjust = 0.5, nudge_y = 0.03) +
  geom_path()

ggplotly(plt)

# SCIE 논문 건수
transposed <- read_excel(t_names[6]) %>% replace(is.na(.), 0) %>% melt(id.vars = c('Pair'))
transposed <- cbind(paste(transposed$Pair, "-", transposed$variable, sep = ""), transposed)
colnames(transposed) <- c('Code', 'Year', 'Pair', tablenames[6])
transposed <- transposed[,-2]
transposed <- transposed[,-2]

p <- merge(merged, transposed, by = "Code", all = TRUE)
plt <- p %>%
  ggplot(aes(x = Projects, y = SCIE, group = Pair, color = Pair, dlabel = Year)) + 
  geom_point(size = 2) +
  ggtitle("SCIE 논문 건수") + xlab("정부연구비(normalized)") + ylab("SCIE 논문 건수(normalized)") +
  geom_text(data = filter(p, Year == 2021), aes(label = Pair), size = 5, nudge_x = 0.06) +
  geom_text(data = p, aes(label = Year), hjust = 1, vjust = 0.5, nudge_y = 0.03) +
  geom_path()

ggplotly(plt)
