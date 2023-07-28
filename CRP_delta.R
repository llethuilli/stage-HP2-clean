
# ============================ IMPORT LIBRARIES ================================

library(readxl)
library(MASS) # load before dplyr to not have 'select' function hidden
library(tidyr)
library(dplyr)
library(ggplot2)
library(stringr)
library(gtsummary)
library(lmerTest)
library(lubridate)
library(labelled)

# ============================= IMPORT DATA ====================================

# set working directory
setwd("D:/Documents_D/INSA/4A/stage/etude_CRP_Lea_Lethuillier")

# path to export figures / tables
output_path = 'output_delta/'

#data = read_excel("data/2023_03_23_ESADA_CLEAN_EXPORT_V0_TO_5_WITH_SCORES.xlsx") #59,907
load("data/bd_esada_initial.RData")

# load correctly the variable Cause.Of.Fatality (as a text variable)
cause = read_excel("data/2023_03_23_ESADA_CLEAN_EXPORT_V0_TO_5_WITH_SCORES.xlsx",
                   sheet = 3, col_types = c('numeric','numeric','text'))

data = data %>% select(-Cause.Of.Fatality) %>%
  left_join(cause, by = c('Patient.ID', 'visit'))

# Use.h.day. is loaded in the first dataframe as TRUE/FALSE
# load it correctly (as numeric)
observance = read_excel("data/2023_03_23_ESADA_CLEAN_EXPORT_V0_TO_5_WITH_SCORES.xlsx",
                         sheet = 2, col_types = c('numeric','numeric','numeric'))

data = data %>% select(-Use.h.day.) %>% 
  left_join(observance, by = c('Patient.ID', 'visit'))

# load correctly PAP treatment start
start_treatment = read_excel("data/2023_03_23_ESADA_CLEAN_EXPORT_V0_TO_5_WITH_SCORES.xlsx",
                             sheet = 4, col_types = c('numeric','numeric', 'text'))
data = data %>% select(- PAP.Treatment.start) %>%
  left_join(start_treatment, by = c('Patient.ID', 'visit'))


# ======================= SELCTION OF THE POPULATION ===========================

# visits of patients with correct values of CRP
data2 = data %>%
  drop_na(C.Reactive.Protein.mg.L.) %>% # 25,189
  filter(C.Reactive.Protein.mg.L. > 0) %>% # 24,970
  filter(C.Reactive.Protein.mg.L. <= 20) # 23,994
length(unique(data2$Patient.ID)) #19,267


# remove patients with treatment ATC H02 in one of their visits
load("D:/Documents_D/INSA/4A/stage/ATC_patients.RData")
#ATC = read_excel("ATC_patients.xlsx", sheet = 3)
ATC <- ATC %>% filter(Patient.ID %in% data2$Patient.ID) %>%
  select(Patient.ID, visit, H02)

data2 = data2 %>% left_join(ATC, by = c('Patient.ID', 'visit')) 

patients_H02 = data2 %>% filter(H02 == 1) %>% filter(visit == 0 | visit == 1)
data2 = data2 %>% filter(! Patient.ID %in% patients_H02$Patient.ID) #23,680
length(unique(data2$Patient.ID)) #19,049


# remove patients with cancer
data2 = data2 %>%
  mutate(cancer =
           ifelse(grepl('cancer', data2$Other, ignore.case = TRUE)|
                    grepl('cancer', data2$Pulmonary...Other, ignore.case = TRUE)|
                    grepl('cancer', data2$Metabolic...Other, ignore.case = TRUE)|
                    grepl('cancer', data2$Cause.Of.Fatality, ignore.case = TRUE), 1, 0))

patients_cancer = data2 %>% filter(cancer == 1)
data2 = data2 %>% filter(! Patient.ID %in% patients_cancer$Patient.ID) # 23,626
length(unique(data2$Patient.ID)) # 19,003


# patients with visit 0 and visit 1 and weight and observance
patients0 = data2 %>%
  filter(visit == 0 & !is.na(Weight) & !is.na(AHI)) #18,315

patients1 = data2 %>%
  filter(visit == 1  & !is.na(Use.h.day.) & PAP...yes.no == 1 & !is.na(Weight)) #2,497

data2 = data2 %>% 
  filter(Patient.ID %in% patients0$Patient.ID & Patient.ID %in% patients1$Patient.ID) %>%
  filter(visit == 0 | visit == 1) #4,602

nb_patients = length(unique(data2$Patient.ID)) #2,306

# ======================== CREATION OF NEW VARIABLES ===========================

#create an categorical variable for observance
data3 <- data2 %>%
  mutate(observance = ifelse(Use.h.day. < 4, '<4', '>=4'))

data3 = data3 %>%
  mutate(observance = factor(observance, levels = c('<4', '>=4')))

# change level names of AHI_class
data3$AHI_class <- 
  recode_factor(data3$AHI_class, 'No SA' = 'No OSA', 'Mild SA' = 'Mild OSA',
                'Moderate SA' = 'Moderate OSA', 'Severe SA' = 'Severe OSA')


# creation of a variable year of the visit
data3 = data3 %>% mutate(Visit.Date = as.POSIXct(Visit.Date)) %>%
  mutate(Visit.Year = as.numeric(format(Visit.Date, "%Y")))

# treatment start as POSIXct instead of character
data3 <- data3 %>% mutate(PAP.Treatment.start = as.POSIXct(PAP.Treatment.start))

# gender as numeric instead of character
data3 <- data3 %>% mutate(Sex_Male = ifelse(Gender == 'Male', 1, 0))

# create a smoking variable
data3 <- data3 %>% mutate(Smoking = ifelse(Smoking.Not.Set, NA,
                                           ifelse(Smoking...Yes, 1, 0)))

# merge variables of diabetes type I and II into only one variable diabetes
data3 = data3 %>% 
  mutate(Metabolic.Diabetes = ifelse(
    Metabolic.Diabetes...Non.insulin.dependent == 1 |
      Metabolic.Diabetes...Insulin.Dependent == 1, 1, 0))

# create a variable for hypertension (systemic and pulmonary)
# and a variable for all other CV comorbidities 
data3 <- data3 %>% mutate(
  CV.hypertension = ifelse(CV.Systemic.Hypertension | CV.Pulmonary.Hypertension, 1, 0),
  CV.Other = ifelse(CV.Valvular.Heart.Disease | Atrial.Fibrillation |
                      CV.Other.Cerebrovascular.Disease | CV.Arrhythmia |
                      CV.Tachycardia | CV.Vein.Insufficiency | CV.Cardiomyopathy |
                      CV.Coagulation.Disorder | CV.Other.Yes.No, 1, 0))

# list of all variables
quanti_var = c('Age','Weight', 'AHI', 'C.Reactive.Protein.mg.L.', 'ESS', 'Visit.Year')
binary_var = c('Sex_Male', 'Smoking')
cat_var = c('AHI_class', 'observance')

comorbidities = c('CV.Left.Ventricular.Hypertrophy', 'CV.hypertension',
                  'CV.Ischemic.Heart.Disease', 'CV.TIA.or.stroke', 
                  'CV.Status.Post.Myocardial.Infarction', 'CV.Cardiac.failure',
                  'CV.Other','Metabolic.Diabetes', 'Pulmonary...COPD',
                  'Other...Neurological.Disease', 'Other...Psychiatric.Disease',
                  'Other...Inflammatory.Disease')


# ============================== DATA-MANAGEMENT ===============================

# replace missing values of comorbidities by 0
data4 <- mutate_at(data3, comorbidities, ~replace_na(.,0))

# replace missing values of smoking by 0
data4 <- mutate_at(data4, 'Smoking', ~replace_na(.,0))

# replace missing values of ESS by the median
data4 <- mutate_at(data4, c("ESS"), ~replace_na(., median(., na.rm = TRUE)))

# put categorical variables of interest as factor
data4 <- data4 %>% mutate_at(c(cat_var, binary_var, comorbidities, 'Site', 'Patient.ID'), ~as.factor(.))

# change the reference level of AHI class to have 'No SA' as reference
data4 <- data4 %>% mutate(AHI_class = relevel(AHI_class, ref = "No OSA"))


############################ FIRST METHOD (delta CRP)###########################

# ===================== CREATION OF A DF : one row per patient =================

# data.frame with only baseline visits (one row per patient)
df_visit0 = data4 %>% filter(visit == 0) %>%
  select(Patient.ID, C.Reactive.Protein.mg.L., AHI_class, Age, ESS, Weight, Smoking,
         Sex_Male, Metabolic.Diabetes, CV.Left.Ventricular.Hypertrophy,
         CV.hypertension, CV.Ischemic.Heart.Disease, CV.TIA.or.stroke,
         CV.Status.Post.Myocardial.Infarction, CV.Cardiac.failure, CV.Other,
         Metabolic.Diabetes, Pulmonary...COPD, Other...Neurological.Disease,
         Other...Psychiatric.Disease, Other...Inflammatory.Disease, Site, Visit.Year)


# data.frame with additional information from visit 1
infos_visit1 = data4 %>% arrange(visit) %>% group_by(Patient.ID) %>%
  summarise(delta_CRP = (C.Reactive.Protein.mg.L.[2] - C.Reactive.Protein.mg.L.[1]),
            days = as.numeric(round(Visit.Date[2] - PAP.Treatment.start[2])), #in days
            delta_weight = Weight[2]-Weight[1],
            observance = observance[2],
            Use.h.day. = Use.h.day.[2])

infos_visit1 = infos_visit1 %>% mutate(observance_binary = ifelse(observance == '<4',0 , 1))

# dataframe with baselie characteristics and infos from visit 1 (still one line per patient)
df = df_visit0 %>% left_join(infos_visit1, by = 'Patient.ID')

# remove patients with visit 1 before visit
patients_pb_dates = df %>% filter(days < 0) # 15 patients
df = df %>% filter(! Patient.ID %in% patients_pb_dates$Patient.ID) #2286

all_var = names(df)
all_var = all_var[-c(which(all_var == "Patient.ID"), which(all_var == "Site"),
                     which(all_var == "Visit.Year"))]

quanti_var = c('C.Reactive.Protein.mg.L.', 'Age', 'ESS', 'Weight', 'delta_CRP',
               'days', 'delta_weight')

cat_var = c('AHI_class','observance')

binary_var = c('Smoking', 'Sex_Male', 'Metabolic.Diabetes', 'CV.Left.Ventricular.Hypertrophy',
               'CV.hypertension', 'CV.Ischemic.Heart.Disease', 'CV.TIA.or.stroke',
               'CV.Cardiac.failure', 'Pulmonary...COPD', 'Other...Neurological.Disease',
               'Other...Psychiatric.Disease', 'Other...Inflammatory.Disease')


# add labels to the dataframe
my_labels = c(
  'Patient.ID' = 'Patient ID', 'C.Reactive.Protein.mg.L.' = 'Baseline C Reactive Protein (mg/L)',
  'AHI_class' = 'OSA severity', 'Age' = 'Age', 'ESS' = 'ESS', 'Baseline weight' = 'Weight (kg)',
  'Smoking' = 'Smoking', 'Sex_Male' = 'Sex male', 'Metabolic.Diabetes' = 'Diabetes',
  'CV.Left.Ventricular.Hypertrophy' = 'Left ventricular hypertrophy',
  'CV.hypertension' = 'Hypertension', 'CV.Ischemic.Heart.Disease' = 'Ischemic heart disease',
  'CV.TIA.or.stroke' = 'TIA or stroke',
  'CV.Status.Post.Myocardial.Infarction' = 'Status post myocardial infarction',
  'CV.Cardiac.failure' = 'Cardiac failure', 'CV.Other' = 'Other CV',
  'Pulmonary...COPD' = 'COPD', 'Other...Neurological.Disease' = 'Neurological disease',
  'Other...Psychiatric.Disease' = 'Psychiatric disease',
  'Other...Inflammatory.Disease' = 'Inflammatory disease', 'Site' = 'Site',
  'Visit.Year' = 'Year of the visit', 'delta_CRP' = 'Delta CRP',
  'days' = 'Treatment duration (days)', 'delta_weight' = 'Delta weight',
  'observance' = 'Adherent (>= 4h)', 'Use.h.day' = 'Use per day (hour)',
  'observance_binary' = 'Observance binary')

df <- set_variable_labels(df, .labels = my_labels)

# ============================= DATA DESCRIPTION ===============================

# description with median (IQR) and n(%)
description_after = df %>%
  select(all_of(c(quanti_var, binary_var, cat_var))) %>%
  tbl_summary(
    statistic = list(all_continuous() ~ "{median} ({p25}, {p75})",
                     all_categorical() ~ "{n} ({p}%)" ),
    type = list(all_of(binary_var) ~ "dichotomous"),
    digits = all_continuous() ~ 2,
    missing = 'no')

# number and percent of missing values
description_missing_after = df %>%
  select(all_of(c(quanti_var, binary_var, cat_var))) %>%
  tbl_summary(
    statistic = list(all_continuous() ~ "{N_miss} ({p_miss}%)",
                     all_categorical() ~ "{N_miss} ({p_miss}%)"),
    type = list(all_of(binary_var) ~ "dichotomous"),
    digits = all_continuous() ~ 2,
    missing='no')

# join the two in one dataframe
all_description_after = tbl_merge(
  list(description_after, description_missing_after),
  tab_spanner = c('Data description', 'Missing values')
)

# all_description_after %>%
#   as_flex_table() %>%
#   flextable::save_as_docx(path = paste(output_path, 'description.docx', sep = ''))


# description by observance group
description_by_observance = df %>%
  select(all_of(c(quanti_var, binary_var, cat_var))) %>%
  tbl_summary(
    by = observance,
    statistic = list(all_continuous() ~ "{median} ({p25}, {p75})",
                     all_categorical() ~ "{n} ({p}%)" ),
    type = list(all_of(binary_var) ~ "dichotomous"),
    digits = all_continuous() ~ 2,
    missing = 'no') %>%
  add_p()

# description_by_observance %>%
#   as_flex_table() %>%
#   flextable::save_as_docx(path = paste(
#     output_path, 'description_by_group_observance.docx', sep = '')
#     )

# =================================== GRAPHS ===================================

# histogram
ggplot(df, aes(x = delta_CRP)) +
  geom_histogram(color = "black", fill = "grey90", binwidth = 1) +
  ggtitle('Histogram of delta CRP') +
  xlab('delta CRP (mg/L)') +
  theme_minimal()

# ============================= CREATION OF MODEL ==============================

model = lmer(delta_CRP ~ observance + days + C.Reactive.Protein.mg.L. + delta_weight +
               Weight +  AHI_class + Sex_Male + Age + ESS + Smoking +
               Metabolic.Diabetes + CV.hypertension + CV.Cardiac.failure + 
               Pulmonary...COPD + Other...Inflammatory.Disease + (1|Site),
             data = df)

binary_var = c('observance','Sex_Male','Smoking','Metabolic.Diabetes','CV.hypertension',
               'CV.Cardiac.failure','Pulmonary...COPD','Other...Inflammatory.Disease')

summary_model1 = model %>% tbl_regression(
  show_single_row = all_of(binary_var),
  pvalue_fun = ~ style_pvalue(.x, digits = 2),
  intercept = TRUE
)

#add global p-values
summary_model1 = add_global_p(summary_model1, include = c(AHI_class),
                             keep = TRUE)

# summary_model1 %>%
#   as_flex_table() %>%
#   flextable::save_as_docx(path = paste(output_path, 'model_delta.docx', sep = ''))
# 
# summary_model1 %>%
#   as_kable_extra(format = "latex") %>%
#   readr::write_lines(file = paste(output_path, "deltaCRP.tex", sep = ''))

# study of the residuals
values = residuals(model)
residuals = data.frame(values)
ggplot(residuals, aes(x = values)) +
  geom_histogram(color = "black", fill = "grey90", binwidth = 1) +
  ggtitle('Histogram of residuals of the model') +
  xlab('residuals') +
  theme_minimal()

plot(model, main = 'Residuals vs fitted values', xlab = 'fitted values',
     ylab = 'residuals')

#study of the random effects
ranova(model)


####################### SECOND METHOD: MIXED MODELS ############################

# ===================== CREATION OF A DF : one row per visit ===================

### Creation of a dataframe with visit 0 of patients
# df with visit 0 of patients
df_visit0 = data4 %>% filter(visit == 0) %>% 
  select(Patient.ID, C.Reactive.Protein.mg.L., AHI_class, Age, ESS, Weight, Smoking,
         Sex_Male, Metabolic.Diabetes, CV.Left.Ventricular.Hypertrophy,
         CV.hypertension, CV.Ischemic.Heart.Disease, CV.TIA.or.stroke,
         CV.Status.Post.Myocardial.Infarction, CV.Cardiac.failure, CV.Other,
         Metabolic.Diabetes,Pulmonary...COPD, Other...Neurological.Disease,
         Other...Psychiatric.Disease, Other...Inflammatory.Disease, Site, Visit.Year) %>%
  mutate(days = 0)

# retrieve values of observance from visit 1
values_observance = data4 %>% arrange(visit) %>% group_by(Patient.ID) %>%
  summarise(Patient.ID = Patient.ID[2],
            observance = observance[2])

# add value of observance in the dataframe visit 0
df_visit0 = df_visit0 %>% left_join(values_observance, by = 'Patient.ID')

### Creation of a dataframe with visit 1 of patients
df_visit1 = data4 %>% filter(visit == 1) %>%
  select(Patient.ID, C.Reactive.Protein.mg.L., Weight, observance,
         Age, ESS, Smoking, Sex_Male, Metabolic.Diabetes,
         CV.Left.Ventricular.Hypertrophy, CV.hypertension, CV.Ischemic.Heart.Disease,
         CV.TIA.or.stroke, CV.Status.Post.Myocardial.Infarction, CV.Cardiac.failure,
         CV.Other, Metabolic.Diabetes, Pulmonary...COPD, Other...Neurological.Disease,
         Other...Psychiatric.Disease, Other...Inflammatory.Disease, Site, Visit.Year)

# retrieve values of variables of interest from visit 0
values_visit1 = data4 %>% arrange(visit) %>% group_by(Patient.ID) %>%
  summarise(days = as.numeric(round(Visit.Date[2] - Visit.Date[1])),
            AHI_class = AHI_class[1])

# add those values to the dataframe of visit 1
df_visit1 = df_visit1 %>% left_join(values_visit1, by = 'Patient.ID')


# reorder the columns to have the same order in both dataframes
df_visit1 = df_visit1 %>% select(all_of(names(df_visit0)))
# merge the two dataframes into one
df2 = bind_rows(df_visit0, df_visit1)

# remove patients with negative values for time of treatment
patients_pb_dates = df2 %>% filter(days < 0) # 5 patients
df2 = df2 %>% filter(! Patient.ID %in% patients_pb_dates$Patient.ID) #4592


# add labels to rename the variables in the outputs
my_labels = c(
  'Patient.ID' = 'Patient ID', 'C.Reactive.Protein.mg.L.' = 'C Reactive Protein (mg/L)',
  'AHI_class' = 'OSA severity', 'Age' = 'Age', 'ESS' = 'ESS', 'Weight' = 'Weight (kg)',
  'Smoking' = 'Smoking', 'Sex_Male' = 'Sex male', 'Metabolic.Diabetes' = 'Diabetes',
  'CV.Left.Ventricular.Hypertrophy' = 'Left ventricular hypertrophy',
  'CV.hypertension' = 'Hypertension', 'CV.Ischemic.Heart.Disease' = 'Ischemic heart disease',
  'CV.TIA.or.stroke' = 'TIA or stroke',
  'CV.Status.Post.Myocardial.Infarction' = 'Status post myocardial infarction',
  'CV.Cardiac.failure' = 'Cardiac failure', 'CV.Other' = 'Other CV',
  'Pulmonary...COPD' = 'COPD', 'Other...Neurological.Disease' = 'Neurological disease',
  'Other...Psychiatric.Disease' = 'Psychiatric disease',
  'Other...Inflammatory.Disease' = 'Inflammatory disease', 'Site' = 'Site',
  'Visit.Year' = 'Year of the visit', 'days' = 'Time between visits',
  'observance' = 'Observance')

df2 <- set_variable_labels(df2, .labels = my_labels)

# =========================== CREATION OF A MODEL ==============================

model = lmer(C.Reactive.Protein.mg.L. ~ observance + Weight + AHI_class + days +
                Sex_Male + Age + ESS + Smoking + Metabolic.Diabetes +
                CV.hypertension + CV.Cardiac.failure + Pulmonary...COPD +
                Other...Inflammatory.Disease + observance * days + AHI_class:days +
                (1|Patient.ID) + (1|Site) +(1|Visit.Year) ,
              data = df2)

binary_var= c('Sex_Male', 'Smoking', 'Metabolic.Diabetes', 'CV.hypertension',
              'CV.Cardiac.failure', 'Pulmonary...COPD', 'Other...Inflammatory.Disease')

summary_model = model %>% tbl_regression(
  show_single_row = all_of(binary_var),
  pvalue_fun = ~ style_pvalue(.x, digits = 2)
)

# add global p-values
summary_model = add_global_p(summary_model, include = c(AHI_class, observance),
                             keep = TRUE)

# summary_model %>%
#   as_flex_table() %>%
#   flextable::save_as_docx(path = paste(output_path, 'model.docx', sep = ''))

# study of the residuals
values = residuals(model)
residuals = data.frame(values)
ggplot(residuals, aes(x = values)) +
  geom_histogram(color = "black", fill = "grey90", binwidth = 1) +
  ggtitle('Histogram of residuals of the model') +
  xlab('residuals') +
  theme_minimal()

plot(model, main = 'Residuals vs fitted values', xlab = 'fitted values',
     ylab = 'residuals')

# study of the random effects
ranova(model)
