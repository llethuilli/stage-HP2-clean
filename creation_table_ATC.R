
# ========================= IMPORT LIBRARIES ===================================

library(readxl)
library(tidyr)
library(dplyr)
library(writexl)
library(kableExtra)
library(stringr)

# ============================ IMPORT DATA =====================================

setwd("D:/Documents_D/INSA/4A/stage/etude_CRP_Lea_Lethuillier")

#data = read_excel("data/2023_03_23_ESADA_CLEAN_EXPORT_V0_TO_5_WITH_SCORES.xlsx")
load("data/bd_esada_initial.RData")

data <- data %>% select(Patient.ID, visit, starts_with('ATC'))

# merge all strings of ATC for each patient in a single column
data <- data %>% 
  mutate(ATC = paste(ATC1, ATC2, ATC3, ATC4, ATC5, ATC6, ATC7, ATC8, ATC9,
                     ATC10, ATC11, ATC12, ATC13, ATC14, ATC15, ATC16, ATC17,
                     ATC18, ATC19, ATC20, ATC21, ATC22, ATC23, ATC24))


# There are two levels of classification :
# first level classification = only the letter (ex: A)
# second level classification = the letter and two digits (ex: A01)

# import ATC codes and corresponding classification for the first level classification
category_ATC = read_excel('data/codes_ATC.xlsx', sheet = 1)

# import the ATC codes and corresponding classification for the second level classification
code_ATC = read_excel('data/codes_ATC.xlsx', sheet = 2)


# ======================= SECOND LEVEL CLASSIFICATION ==========================

# creation of a new data.frame with a binary column for each ATC code

list_codes = code_ATC$code_ATC #list of ATC codes
second_level <- data.frame(Patient.ID = data$Patient.ID, visit = data$visit)
# 1 in the column if the patient has the treatment, else 0
for(i in 1:length(list_codes)){
  second_level[list_codes[i]] = ifelse(grepl(list_codes[i], data$ATC), 1, 0) 
}

# ======================== FIRST LEVEL CLASSIFIATION ===========================

list_categories = category_ATC$code_category_ATC #list of ATC categories

# creation of a new data.frame with column per category
# 1 if the patient has at least one treatment of this category, else 0

first_level <- second_level %>% 
  mutate(A = ifelse(A01|A02|A03|A04|A05|A06|A07|A08|A09|A10|A11|A12|A13|A14|A15|A16, 1, 0)) %>%
  mutate(B = ifelse(B01|B02|B03|B05|B06, 1, 0)) %>%
  mutate(C = ifelse(C01|C02|C03|C04|C05|C07|C08|C09|C10, 1, 0)) %>%
  mutate(D = ifelse(D01|D02|D03|D04|D05|D06|D07|D08|D09|D10|D11, 1, 0)) %>%
  mutate(G = ifelse(G01|G02|G03|G04, 1, 0)) %>%
  mutate(H = ifelse(H01|H02|H03|H04|H05, 1, 0)) %>%
  mutate(J = ifelse(J01|J02|J04|J05|J06|J07, 1, 0)) %>%
  mutate(L = ifelse(L01|L02|L03|L04, 1, 0)) %>%
  mutate(M = ifelse(M01|M02|M03|M04|M05|M09, 1, 0)) %>%
  mutate(N = ifelse(N01|N02|N03|N04|N05|N06|N07, 1, 0)) %>%
  mutate(P = ifelse(P01|P02, 1, 0))%>%
  mutate(R = ifelse(R01|R02|R03|R05|R06|R07, 1, 0)) %>%
  mutate(S = ifelse(S01|S02, 1, 0)) %>%
  mutate(V = ifelse(V01|V03|V04|V06|V07|V08, 1, 0))  %>%
  select(Patient.ID, visit, A, B, C, D, G, H, J, L, M, N, P, R, S, V) #select only the categories

# ========================= DESCRIPTION OF THE TABLES ==========================

# summarize data with count, percentage of missing values and 
# percentage of patients with ATC for each variable

summarizer_binary <- function(data, cols = NULL) {
  res = data %>%
    summarise(across(all_of({{cols}}), list(
      Count = ~n(),
      Missing_Values = ~paste(
        format(length(which(is.na(.)))/length(.)*100, nsmall = 2, digits = 3),
        "%"),
      Percentage_patients_with_ATC = ~paste(
        format(sum(. == 1, na.rm = TRUE)/length(which(!is.na(.)))*100,
               nsmall = 2, digits = 3),"%"
        )
    ), .names = "{col}-{fn}"))
  
  df = as.data.frame(res)
  df_tidy <- df %>% gather(stat, val) %>%
    separate(stat, into = c("Var", "stat"), sep = "-") %>%
    spread(stat, val) %>%
    select(Var, Count, Missing_Values, Percentage_patients_with_ATC)
}

#table of description for second level classification
table_second_level = summarizer_binary(second_level, cols = list_codes)

#add text description (in english or french)
table_second_level_english = cbind(table_second_level, 
                                   Classification = code_ATC$english_classification)
table_second_level_french = cbind(table_second_level,
                                  Classification = code_ATC$french_classification)

#table of description for first level classification
table_first_level = summarizer_binary(first_level, cols = list_categories)
#add text description (in english or french)
table_first_level_english=cbind(table_first_level,
                                Classification = category_ATC$english_classification)
table_first_level_french=cbind(table_first_level,
                               Classification = category_ATC$french_classification)


# export table to html
print_table <- function(df, caption = '') {
  df %>%
    kbl(caption = {{caption}}) %>%
    kable_classic(full_width = T) %>%
    kable_styling(font_size = 13)
}

# English version
print_table(table_first_level_english, 'First level ATC description')
print_table(table_second_level_english, 'Second level ATC description')

# French version 
# print_table(table_first_level_french)
# print_table(table_second_level_french)

# =========================  EXPORT TO EXCEL FILE ==============================

# creation of an excel file with the data.frames for first and second level classification
# and the description of the classification in French and English

sheets <- list("first_level_data" = first_level, 
               "first_level_classification" = category_ATC,
               "second_level_data" = second_level,
               "second_level_classification" = code_ATC)

write_xlsx(sheets, "data/ATC_patients.xlsx")
