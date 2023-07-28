
# ============================ IMPORT LIBRARIES ================================

library(readxl)
library(MASS) # load before dplyr to not have 'select' function hidden
library(tidyr)
library(dplyr)
library(ggplot2)
library(kableExtra)
library(tidyverse)
library(lmerTest)
library(gridExtra)
library(broom)
library(ipw)
library(tableone)
library(gt)
library(survey)
library(broom)
library(broom.mixed)
library(gtsummary)
library(flextable)
library(ordinal)
library(ggdag)
library(moments)
library(boot)
library(labelled)
library(performance)
library(docstring)
library(writexl)

# ============================= IMPORT DATA ====================================

# set working directory
setwd("D:/Documents_D/INSA/4A/stage/etude_CRP_Lea_Lethuillier")

# path to export the tables
output_path = 'output_baseline/'

# data = read_excel("data/2023_03_23_ESADA_CLEAN_EXPORT_V0_TO_5_WITH_SCORES.xlsx") #59,907
load("data/bd_esada_initial.RData")

# load correctly the variable Cause.Of.Fatality (as a text variable)
cause = read_excel("data/2023_03_23_ESADA_CLEAN_EXPORT_V0_TO_5_WITH_SCORES.xlsx",
                   sheet = 3, col_types = c('numeric','numeric','text'))

data = data %>% select(-Cause.Of.Fatality) %>%
  left_join(cause, by = c('Patient.ID', 'visit'))


# ======================= SELECTION OF THE POPULATION ==========================

# keep only visit 0, with values of AHI and CRP, and 0 < CRP <= 20mg/L
data2 = data %>% filter(visit == 0) %>% # 36,308 patients
  drop_na(C.Reactive.Protein.mg.L.) %>% # 19,827
  filter(C.Reactive.Protein.mg.L. > 0) %>% # 19,641
  drop_na(AHI) %>% # 19,447
  filter(C.Reactive.Protein.mg.L. <= 20) # 18,678

# remove patients with treatment ATC H02
load("D:/Documents_D/INSA/4A/stage/ATC_patients.RData")
#ATC = read_excel("ATC_patients.xlsx", sheet = 3)
ATC <- ATC %>% filter(Patient.ID %in% data2$Patient.ID) %>%
  filter(visit == 0) %>%
  select(Patient.ID, visit, H02)

data2 = data2 %>% left_join(ATC, by = c('Patient.ID','visit')) %>%
  filter(H02 == 0) # 18,487

# remove patients with cancer
data2 = data2 %>%
  mutate(cancer =
           ifelse(grepl('cancer', data2$Other, ignore.case = TRUE)|
                    grepl('cancer', data2$Pulmonary...Other, ignore.case = TRUE)|
                    grepl('cancer', data2$Metabolic...Other, ignore.case = TRUE)|
                    grepl('cancer', data2$Cause.Of.Fatality, ignore.case = TRUE), 1, 0))%>%
  filter(cancer == 0) #18,445


# ========================= CREATION OF NEW VARIABLES ==========================

# create an categorical variable for BMI
data3 <- data2 %>% 
  mutate(BMI_class = ifelse(BMI < 25, '<25',
                            ifelse(BMI <= 30, '25-30',
                                   ifelse(BMI <= 35, '30-35', '>35'))))

# change level names of AHI_class
data3$AHI_class <- 
  recode_factor(data3$AHI_class, 'No SA' = 'No OSA', 'Mild SA' = 'Mild OSA',
                'Moderate SA' = 'Moderate OSA', 'Severe SA' = 'Severe OSA')

# create ordinal variable for AHI class
data3 = data3 %>% 
  mutate(AHI_ordinal = factor(AHI_class, order = TRUE,
                              levels = c('No OSA', 'Mild OSA', 'Moderate OSA',
                                         'Severe OSA')))

# creation of a variable year of the visit
data3 = data3 %>% mutate(Visit.Date = as.POSIXct(Visit.Date)) %>%
  mutate(Visit.Year = as.numeric(format(Visit.Date, "%Y")))

# gender as numeric instead of character
data3 <- data3 %>% mutate(Sex_Male = ifelse(Gender == 'Male', 1, 0))

# Change the name of site to have the city only
data3$Site = str_replace_all(
  data3$Site, "Warsaw Institute of Tuberculosis and Lung Diseases", "Warsaw")

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
quanti_var = c('Age', 'BMI', 'AHI', 'C.Reactive.Protein.mg.L.', 'ESS', 'Visit.Year')
binary_var = c('Sex_Male', 'Smoking')
cat_var = c('AHI_class','BMI_class')

comorbidities = c('CV.Left.Ventricular.Hypertrophy', 'CV.hypertension',
                  'CV.Ischemic.Heart.Disease', 'CV.TIA.or.stroke', 
                  'CV.Status.Post.Myocardial.Infarction', 'CV.Cardiac.failure',
                  'CV.Other','Metabolic.Diabetes', 'Pulmonary...COPD',
                  'Other...Neurological.Disease', 'Other...Psychiatric.Disease',
                  'Other...Inflammatory.Disease')


# ============== GENERAL DESCRIPTION (before data-management) ==================

# description with median (IQR) and n(%)
description_before = data3 %>%
  select(all_of(c(quanti_var, binary_var, cat_var, comorbidities))) %>%
  tbl_summary(
    statistic = list(all_continuous() ~ "{median} ({p25}, {p75})",
                      all_categorical() ~ "{n} ({p}%)" ),
    digits = all_continuous() ~ 2,
    missing = 'no')

# number and percent of missing values
description_missing_before = data3 %>%
  select(all_of(c(quanti_var, binary_var, cat_var, comorbidities))) %>%
  tbl_summary(
    statistic = list(all_continuous() ~ "{N_miss} ({p_miss}%)",
                     all_categorical() ~ "{N_miss} ({p_miss}%)"),
    digits = all_continuous() ~ 2,
    missing='no')

# join the two in one dataframe
all_description_before = tbl_merge(
  list(description_before, description_missing_before),
  tab_spanner = c('Data description', 'Missing values')
)

# all_description_before %>%
#   as_flex_table() %>%
#   flextable::save_as_docx(path = paste(output_path, 'final_summary.docx', sep = ''))
# 
# all_description_before %>%
#   as_kable_extra(format = "latex") %>%
#   readr::write_lines(file = "latex_table.tex")

# ============================== DATA-MANAGEMENT ===============================

# if BMI > 70, data of height, weight and BMI are considered missing
data3 = data3 %>% mutate(BMI = Weight/ (Height/100)^2)
abnormal_values = which(data3$BMI > 70) #list of patients with BMI > 70
data3[abnormal_values, ]$BMI = NA
data3[abnormal_values, ]$Height = NA
data3[abnormal_values, ]$Weight = NA

# replace missing values of comorbidities by 0
data4 <- mutate_at(data3, comorbidities, ~replace_na(.,0))

# replace missing values of smoking by 0
data4 <- mutate_at(data4, 'Smoking', ~replace_na(.,0))

# replace missing values of height and weight by the median
data4 <- mutate_at(data4, c("Height", "Weight"), ~replace_na(., median(., na.rm = TRUE)))

# compute the missing BMI with the new values of height and weight
data4$BMI[is.na(data4$BMI)] <- 
  data4$Weight[is.na(data4$BMI)]/((data4$Height[is.na(data4$BMI)]/100)^2)

# create BMI classes with the new values of BMI
data4 <- data4 %>% 
  mutate(BMI_class = ifelse(BMI < 25, '<25',
                            ifelse(BMI <= 30, '25-30',
                                   ifelse(BMI <= 35, '30-35', '>35')))) %>%
  mutate(BMI_class = factor(BMI_class, levels = c("<25", "25-30", "30-35", ">35")))

# replace missing values of ESS by the median
data4 <- mutate_at(data4, c("ESS"), ~replace_na(., median(., na.rm = TRUE)))

## put categorical variables of interest as factor
data4 <- data4 %>% mutate_at(c(cat_var, binary_var, comorbidities, 'Site', 'Patient.ID'),
                             ~as.factor(.))

# change the reference level of AHI class to have 'No SA' as reference
data4 <- data4 %>% mutate(AHI_class = relevel(AHI_class, ref = "No OSA"))


# ========================= CREATE FINAL DATAFRAME =============================

# create a data-frame with the variables of interest only
final_df = data4 %>%
  select(Patient.ID, C.Reactive.Protein.mg.L., AHI_class, AHI_ordinal, Age, ESS,
         BMI_class, Smoking,Sex_Male, Metabolic.Diabetes, 
         CV.Left.Ventricular.Hypertrophy, CV.hypertension, CV.Ischemic.Heart.Disease, 
         CV.TIA.or.stroke, CV.Status.Post.Myocardial.Infarction,
         CV.Cardiac.failure, CV.Other, Metabolic.Diabetes,Pulmonary...COPD,
         Other...Neurological.Disease, Other...Psychiatric.Disease,
         Other...Inflammatory.Disease, Site, Visit.Year)

all_var = names(final_df)
all_var = all_var[-c(which(all_var == "Patient.ID"), which(all_var == "Site"),
                     which(all_var == "Visit.Year"))]

cat_var = c('AHI_class', 'BMI_class')
binary_var = c('Smoking', 'Sex_Male', 'Metabolic.Diabetes', 'CV.Left.Ventricular.Hypertrophy',
             'CV.hypertension', 'CV.Ischemic.Heart.Disease',
             'CV.TIA.or.stroke', 'CV.Status.Post.Myocardial.Infarction',
             'CV.Cardiac.failure', 'CV.Other', 'Pulmonary...COPD',
             'Other...Neurological.Disease', 'Other...Psychiatric.Disease',
             'Other...Inflammatory.Disease')

# add labels to rename the variables in the outputs
my_labels = c('Patient ID', 'C Reactive Protein (mg/L)', 'OSA severity', 
              'AHI ordinal', 'Age', 'ESS', 'BMI', 'Smoking', 'Sex Male',
              'Diabetes', 'Left ventricular hypertrophy', 'Hypertension', 
              'Ischemic heart disease', 'TIA or stroke', 
              'Status post myocardial infarction', 'Cardiac failure', 
              'Other cardiovascular comorbidities', 'COPD', 'Neurological disease',
              'Psychiatric disease', 'Inflammatory disease', 'Site', 'Year of visit')

final_df <- set_variable_labels(final_df, .labels = my_labels)

# export the dataframe
#write_xlsx(final_df, 'data/final_dataframe_baseline.xlsx')

# ==================== DESCRIPTION (after data-management) =====================

# comparison before and after data-management

description_imputed_var <- function(data, cols = NULL){
  #' Description of imputed variables
  #' 
  #' Creates a description of the selected variables
  #' @param data the dataframe containing the variables of interest
  #' @param cols variables from the dataframe to describe
  #' @return a tbl_summary object
  data %>%
    select(all_of(cols)) %>%
    tbl_summary(
      statistic = list( all_continuous() ~ "{median} ({p25}, {p75})",
                        all_categorical() ~ "{n} ({p}%)" ),
      type = all_categorical()~'dichotomous',
      digits = all_continuous() ~ 2,
      missing = 'no')
}

imputed_var = c('Height', 'Weight', 'BMI', 'ESS', 'Smoking', comorbidities)
before = description_imputed_var(data3, imputed_var)
after = description_imputed_var(data4, imputed_var)

comparison_before_after = tbl_merge(
  list(before, after),
  tab_spanner = c('Before', 'After')
)

# comparison_before_after %>%
#   as_flex_table() %>%
#   flextable::save_as_docx(path = paste(output_path, 'comparison_imputed_vals.docx', sep = ''))
# 
# comparison_before_after %>%
#   as_kable_extra(format = "latex") %>%
#   readr::write_lines(file = paste(output_path, "comparison imputed_vals.tex", sep = ''))


# description of all variables after data-management
description_after = final_df %>%
  select(-Site, -Visit.Year, -Patient.ID, - AHI_ordinal) %>%
  tbl_summary(type = list(all_of(binary_var) ~ "dichotomous"))

description_group_after = final_df %>%
  select(-Site, -Visit.Year, -Patient.ID, - AHI_ordinal) %>%
  tbl_summary(by = AHI_class, type = list(all_of(binary_var) ~ "dichotomous")) %>%
  add_p()


# description_group_after %>% 
#   as_flex_table() %>%
#   flextable::save_as_docx(
#     path = paste(output_path, 'description_group_after_datamanagement.docx', sep = '')
#     )
# 
# description_group_after %>%
#   as_kable_extra(format = "latex") %>%
#   readr::write_lines(file = paste(output_path, "description_group_after.tex", sep = ''))


# ================================== GRAPHS ====================================

# histogram of CRP
ggplot(final_df, aes(x = C.Reactive.Protein.mg.L.)) + 
  geom_histogram(color = "black", fill = "grey90", binwidth = 1) +
  ggtitle('Histogram of CRP') +
  xlab('CRP (mg/L)') +
  theme_minimal()

# scatter plot and densities x + y
scatter = ggplot(data4, aes(x = AHI, y = C.Reactive.Protein.mg.L., color = Sex_Male)) +
  geom_point(alpha = 0.5) +
  geom_smooth(aes(group = Sex_Male, fill = Sex_Male), method = lm, fullrange = TRUE,
              linewidth = 0.7, colour = 'grey30') +
  scale_color_discrete(name = "Sex", labels = c("Women", "Men")) +
  scale_fill_discrete(name = "Sex", labels = c("Women", "Men")) +
  ylab('C Reactive Protein (mg/L)') +
  theme_minimal()

xdensity = ggplot(data4, aes(AHI, fill = Sex_Male)) + 
  geom_density(alpha = .5) + 
  theme_minimal() +
  theme(legend.position = "none")
 
ydensity = ggplot(data4, aes(C.Reactive.Protein.mg.L., fill = Sex_Male)) + 
  geom_density(alpha = .5) + 
  theme_minimal() +
  theme(legend.position = "none")

blankPlot <- ggplot() + geom_blank(aes(1, 1)) +
  theme(plot.background = element_blank(), 
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(), 
        panel.border = element_blank(),
        panel.background = element_blank(),
        axis.title.x = element_blank(),
        axis.title.y = element_blank(),
        axis.text.x = element_blank(), 
        axis.text.y = element_blank(),
        axis.ticks = element_blank()
  )

grid.arrange(xdensity, blankPlot, scatter, ydensity, 
             ncol = 2, nrow = 2, widths = c(4, 1.5), heights = c(1.5, 4))

# boxplot CRP / OSA severity / sex
ggplot(final_df, aes(x = AHI_class, y = C.Reactive.Protein.mg.L., fill = Sex_Male)) +
  geom_boxplot(alpha = 1) +
  scale_fill_manual(values = c('lightpink', 'skyblue1'), name = "Sex",
                    labels = c("Female", "Male")) +
  xlab('') + ylab('C Reactive Protein (mg / L)') +
  theme_classic() +
  theme(axis.title = element_text(size = 20),
    axis.text.x = element_text(size = 17),
    axis.text.y = element_text(size = 17),
    legend.title = element_text(size = 20),
    legend.text = element_text(size = 16),
    legend.key.size = unit(1, "cm"))


# ========================= UNIVARIATE ANALYSIS ================================

# compute estimate, IC and p-value for each variable
univariate_analysis = final_df %>% select(-AHI_ordinal, -Patient.ID) %>%
  tbl_uvregression(
    method = lmer, #mixed linear model
    y = "C.Reactive.Protein.mg.L.",
    hide_n = TRUE,
    pvalue_fun = ~ style_pvalue(.x, digits = 2),
    formula = "{y} ~ {x}+(1|Site)+(1|Visit.Year)", #random effects
    show_single_row = all_of(binary_var)
  )

univariate_analysis = add_global_p(univariate_analysis,
                                   include = c('BMI_class', 'AHI_class'),
                                   keep=TRUE)


# ============================== CREATION OF A DAG =============================

dag <- dagify(
  CRP ~ AHI_class + Age + BMI_class + ESS + Smoking + Sex + Diabetes +
    CV.comorbidities + neurological.disease + psychiatric.disease + 
    inflammatory.disease + COPD,
  AHI_class ~ Age + BMI_class + ESS + Smoking + Sex + Diabetes + CV.comorbidities +
    psychiatric.disease + inflammatory.disease + COPD,
  labels = c("CRP" = "CRP", "AHI_class" = "AHI \n class", "Age" = "Age", 
             "BMI_class" = "BMI \n class", "ESS" = "ESS", "Smoking" ="Smoking",
             "Sex" = "Sex", "Diabetes" = "Diabetes", "CV.comorbidities" = "CV \n comorbidities",
             "neurological.disease" = "neurological \n disease",
             "psychiatric.disease" = "psychiatric \n disease",
             "inflammatory.disease" = "inflammatory \n disease",
             "Pulmonary.COPD" = "Pulmonary \n COPD"),
  exposure = "AHI_class",
  outcome = "CRP")

ggdag_status(dag, text = FALSE, use_labels = "label", node_size = 10) +
  guides() +
  theme_dag()


# ============================== COMPUTE WEIGHTS ===============================

# choice of variables to include
covariates = c('Age', 'ESS', 'BMI_class', 'Smoking', 'Sex_Male', 'Metabolic.Diabetes',
             'CV.Left.Ventricular.Hypertrophy', 'CV.hypertension',
             'CV.Ischemic.Heart.Disease','CV.TIA.or.stroke', 'CV.Cardiac.failure',
             'Pulmonary...COPD', 'Other...Neurological.Disease', 
             'Other...Psychiatric.Disease', 'Other...Inflammatory.Disease')

factor_var = c('BMI_class', 'Smoking', 'Sex_Male', 'Metabolic.Diabetes',
             'CV.Left.Ventricular.Hypertrophy', 'CV.hypertension',
             'CV.Ischemic.Heart.Disease', 'CV.TIA.or.stroke', 'CV.Cardiac.failure',
             'Pulmonary...COPD', 'Other...Neurological.Disease', 
             'Other...Psychiatric.Disease', 'Other...Inflammatory.Disease')

binary_var = factor_var[-which(factor_var == 'BMI_class')]

final_df = as.data.frame(final_df)

# computation of stabilized weights
weights <- ipwpoint(
  exposure = AHI_ordinal,
  family = "ordinal",
  link = 'logit',
  numerator = ~1, #stabilized
  denominator = ~  Age + ESS + BMI_class + Sex_Male + Smoking  + 
    Metabolic.Diabetes + CV.Left.Ventricular.Hypertrophy  + 
    CV.Ischemic.Heart.Disease  + CV.hypertension + CV.TIA.or.stroke +
    CV.Cardiac.failure + Pulmonary...COPD + Other...Neurological.Disease +
    Other...Psychiatric.Disease + Other...Inflammatory.Disease,
  data = final_df,
  trunc = .01
)

# add weights to the dataframe
final_df <- final_df %>% 
  mutate(ipw = weights$ipw.weights)%>%
  mutate(ipw_trunc = weights$weights.trunc) 

# see the coefficients of the weight model
weights_model = weights$den.mod %>% tbl_regression(
    exponentiate = TRUE,
    show_single_row=all_of(binary_var)
)

# weights_model %>% 
#   as_flex_table() %>%
#   flextable::save_as_docx(
#     path = paste(output_path, 'coefficients_weights_model.docx', sep = '')
#     )
# 
# weights_model %>%
#   as_kable_extra(format = "latex") %>%
#   readr::write_lines(file = paste(output_path, "coefficients_weights_model.tex", sep = ''))


# description of weights
summary(weights$ipw.weights)
ipwplot(weights$ipw.weights, logscale = FALSE)

# distribution of weights by group
weights_by_group = final_df %>%
  group_by(AHI_ordinal) %>%
  summarise_at(vars(ipw), list(mean = mean, sd = sd, min = min, max = max)) %>%
  mutate(mean = format(round(mean, 2), digits = 2)) %>%
  mutate(sd = format(round(sd, 2), digits = 2)) %>%
  mutate(min = format(round(min, 2), digits = 2)) %>%
  mutate(max = format(round(max, 2), digits = 2))

export_table <- function(df, name='', param = FALSE) {
  #' Export Table
  #'
  #' Export a data-frame in rtf format
  #' @param df dataframe to export
  #' @param name name of the exported file, must finish by '.rtf'
  #' @param param if true, the rownames of the dataframe are rownames in the exported table
  #' @return a rtf file
  df = as.data.frame(df)
  gtsave(gt(df, rownames_to_stub = param),
         paste(output_path, {{name}}, sep = ''))
}
export_table(weights_by_group,'summary_weights_by_group.rtf')

# plot of weights by group
ggplot(final_df, aes(x = ipw, fill = AHI_class)) +
  geom_density(colour = 'black', alpha = 0.5) + theme_minimal()


# apply weights to the population
weighted_data = svydesign(ids = ~1, data = final_df, weights = ~ipw)

unweighted_table <- CreateTableOne(
  vars = covariates, factorVars = factor_var, strata = "AHI_class",
  data = final_df, test = TRUE)

weighted_table <- svyCreateTableOne(vars = c(covariates), factorVars = factor_var,
                    strata = "AHI_class", data = weighted_data, test = TRUE)

# description of groups after weighting
weighted_table_print = as.data.frame(print(weighted_table, printToggle = FALSE))
weighted_table_print = weighted_table_print %>% select(-test)
rownames(weighted_table_print) = c('n', 'C Reactive Protein (mg/L)', 'Age', 'ESS',
                                   'BMI', '<25', '25-30', '30-35', '>35', 'Smoking',
                                   'Sex Male', 'Diabetes', 'Left ventricular hypertrophy',
                                   'Hypertension', 'Ischemic heart disease', 'TIA or stroke',
                                   'Cardiac failure', 'COPD', 'Neurological disease',
                                   'Psychiatric disease', 'Inflammatory disease')
weighted_table_print = tibble::rownames_to_column(weighted_table_print, 'variable')
export_table(weighted_table_print, 'weight_table.rtf')


# plot SMD for weighted and unweighted
plot_SMD <- function(unweighted_table, weighted_table){
  # extract the name of variables and SMD for weighted and unweighted
  dataPlot <- data.frame(variable = rownames(ExtractSmd(unweighted_table, varLabels = TRUE)),
                         Unweighted = as.numeric(ExtractSmd(unweighted_table)[,1]),
                         Weighted = as.numeric(ExtractSmd(weighted_table)[,1]))
  
  #long format
  dataPlotLong <- dataPlot %>% pivot_longer(cols = c('Unweighted', 'Weighted'),
                                            names_to = 'Method',
                                            values_to = 'SMD')
  
  # Order variable names by magnitude of SMD
  varNames <- as.character(dataPlot$variable)[order(dataPlot$Unweighted)]
  
  # Order factor levels in the same order
  dataPlotLong$variable <- factor(dataPlotLong$variable, levels = varNames)
  
  # Plot using ggplot2
  ggplot(data = dataPlotLong,
         mapping = aes(x = variable, y = SMD, group = Method, color = Method)) +
    geom_point(size = 3) +
    scale_color_manual(values = c('sienna2', 'goldenrod1')) +
    scale_y_continuous(breaks = seq(0, 0.6, by = 0.1)) +
    xlab('Variable') + labs(color = "")+
    geom_hline(yintercept = 0.1, color = "grey30", linewidth = 0.1) +
    coord_flip() +
    theme_light() +
    theme(axis.title = element_text(size = 20),
          axis.text.x = element_text(size = 15),axis.text.y = element_text(size = 16),
          legend.text = element_text(size = 16),
          legend.position= 'bottom')
}

plot_SMD(unweighted_table, weighted_table)


# ============================= WEIGHTED REGRESSION ============================

model <- lmer(C.Reactive.Protein.mg.L. ~ AHI_class + Sex_Male + Age + ESS + 
                BMI_class + Smoking + Metabolic.Diabetes + 
                CV.Left.Ventricular.Hypertrophy + CV.hypertension +
                CV.TIA.or.stroke + CV.Ischemic.Heart.Disease + CV.Cardiac.failure +
                Pulmonary...COPD + Other...Neurological.Disease +
                Other...Psychiatric.Disease + Other...Inflammatory.Disease +
                (1|Site) + (1|Visit.Year), 
               data = final_df, weights = ipw)

summary_model = model %>% tbl_regression(
  add_estimate_to_reference_rows = FALSE,
  show_single_row = all_of(binary_var),
  pvalue_fun = ~ style_pvalue(.x, digits = 2)
)

#add global p-values
summary_model = add_global_p(summary_model, include = c('BMI_class', 'AHI_class'),
                             keep = TRUE)

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
icc(model)
icc(model, by_group = TRUE)
ranova(model)


# ------------- table with univariable + multivariable analysis ----------------

final_summary = tbl_merge(
  list(univariate_analysis, summary_model),
  tab_spanner = c('Univariate analysis', 'Multivariate analysis')
)

# final_summary %>%
#   as_flex_table() %>%
#   flextable::save_as_docx(path = paste(output_path, 'final_summary.docx', sep = ''))
# 
# final_summary %>%
#   as_kable_extra(format = "latex") %>%
#   readr::write_lines(file = paste(output_path, "final_summary.tex", sep = ''))


# --------------------------- bootstrap for SE estimates -----------------------

# function to compute the effect estimate of each variable
compute_estimate <- function(formula, data, i) {
  d <- data[i,] #allows boot to select sample
  
  # compute weights
  weights <- ipwpoint(
    exposure = AHI_ordinal,
    family = "ordinal",
    link = 'logit',
    numerator = ~1, #stabilized
    denominator = ~  Age + ESS + BMI_class + Sex_Male + Smoking  + 
      Metabolic.Diabetes + CV.Left.Ventricular.Hypertrophy  + 
      CV.Ischemic.Heart.Disease  + CV.hypertension +CV.TIA.or.stroke+
      CV.Cardiac.failure +
      Pulmonary...COPD + Other...Neurological.Disease +
      Other...Psychiatric.Disease + Other...Inflammatory.Disease,
    data = d,
    trunc = .01
  )
  
  # add weights to the dataframe
  d <- d %>% 
    mutate(ipw = weights$ipw.weights)%>%
    mutate(ipw_trunc = weights$weights.trunc) 
  
  #create the model with truncated weights
  fit <- lmer(formula, data = d, weights = ipw)
  
  #return a vector with the value of the estimates for each variable
  return(as.data.frame(summary(fit)$coefficients)$Estimate)
}

reps_model <- boot(data = final_df, statistic = compute_estimate, R=1000,
            formula = C.Reactive.Protein.mg.L. ~ AHI_class + Sex_Male + Age + 
              ESS + BMI_class + Smoking + Metabolic.Diabetes + 
              CV.Left.Ventricular.Hypertrophy + CV.hypertension +
              CV.TIA.or.stroke + CV.Ischemic.Heart.Disease + CV.Cardiac.failure +
              Pulmonary...COPD + Other...Neurological.Disease +
              Other...Psychiatric.Disease + Other...Inflammatory.Disease +
              (1|Site) + (1|Visit.Year))


compute_res_bootstrap <- function(reps, model){
  
  # retrieve estimate value and create a data-frame
  bootstrap_estimate = reps$t0
  bootstrap = as.data.frame(bootstrap_estimate)
  
  # create a list with all values for each estimate
  list_index = as.list(1:length(reps$t0)) # list of numbers from 1 to nb of estimates
  
  boostrap_values <- function(bootstrap_res, index){
    bootstrap_res$t[,index]
  }
  list_values = lapply(list_index, boostrap_values, bootstrap_res = reps)
  
  # compute SE for each estimate
  std.error = sapply(list_values, sd)
  bootstrap <- bootstrap %>% mutate(SE = std.error)
  
  # create a dataframe with name, estimate, SE and pvalue
  res_bootstrap = as.data.frame(tidy(model)) %>%
    filter(effect == 'fixed') %>%
    select(term, estimate, std.error, p.value) %>%
    mutate (std.error = bootstrap$SE)
  
  return(res_bootstrap)
}

# The line under takes time to compute. The result can directly be loaded from
# the environment res_bootstrap

#res_bootstrap = compute_res_bootstrap(reps_model, model)
load('res_bootstrap.RData')


# Create a summary of the model with bootstrap CI
create_summary <- function(df){
  df = df %>% filter (term != '(Intercept)') %>%
    mutate(
      IC = paste(format(round(estimate - 1.96 * std.error, 2), nsmall = 2), '; ',
          format(round(estimate + 1.96 * std.error, 2), nsmall = 2), sep = '')
      ) %>%
    mutate(estimate = format(round(estimate, 2), digits = 2)) %>%
    mutate(p.value = ifelse(
      p.value < 0.001, '<0.001', format(round(p.value, 3), digits = 3))
    ) %>%
    select(term, estimate, IC, p.value)
  
  return(df)
}

summary_with_bootstrap = create_summary(res_bootstrap)


# Create a graph of coefficients with bootstrap CI

plot_fixed_coefficients <- function(df, my_labels){
  df = df %>%
    filter(term != '(Intercept)') %>%
    mutate(min = estimate - 1.96 * std.error) %>%
    mutate(max = estimate + 1.96 * std.error) %>%
    mutate(significant = as.factor(ifelse(min < 0 & max > 0, 0, 1)))
    
  # Order variable names by magnitude of estimate
  varNames <- as.character(df$term)[order(df$estimate)]
  
  # Order factor levels in the same order
  df$term <- factor(df$term, levels = varNames)
  
  # Create the graph with a different color if significant 
  df %>% 
    ggplot(mapping = aes(x = estimate, y = term, xmin = min, xmax = max, 
                         color = significant)) +
    geom_pointrange(shape = 19) +
    scale_colour_manual(values = c("grey70", "black")) +
    geom_vline(xintercept = 0, color = "red", linewidth = 0.5) +
    labs(x = "Estimate", y = "") +
    scale_y_discrete(labels = my_labels) +
    theme_minimal() +
    theme(axis.title = element_text(size = 20),
          axis.text.x = element_text(size = 17),
          axis.text.y = element_text(size = 17),
          legend.position = "none")
}

labels_coefficients = c(
  'AHI_classMild OSA' = 'Mild OSA', 'AHI_classModerate OSA' = 'Moderate OSA',
  'AHI_classSevere OSA' = 'Severe OSA',
  'BMI_class>35' = 'BMI>35', 'BMI_class30-35' = 'BMI 30-35', 'BMI_class25-30' = 'BMI 25-30',
  'Smoking1' = 'Smoking',
  'Sex_Male1' = 'Sex Male',
  'Metabolic.Diabetes1' = 'Diabetes',
  'CV.Left.Ventricular.Hypertrophy1' = 'Left ventricular hypertrophy',
  'CV.hypertension1' = 'Hypertension',
  'CV.Ischemic.Heart.Disease1' = 'Ischemic heart disease',
  'CV.TIA.or.stroke1' = 'TIA or stroke',
  'CV.Status.Post.Myocardial.Infarction1' = 'Status post myocardial infarction',
  'CV.Cardiac.failure1' = 'Cardiac failure',
  'CV.Other1' = 'Other cardiovascular comorbidities',
  'Pulmonary...COPD1' = 'COPD',
  'Other...Neurological.Disease1' = 'Neurological disease',
  'Other...Psychiatric.Disease1' = 'Psychiatric disease',
  'Other...Inflammatory.Disease1' = 'Inflammatory disease'
  )

plot_fixed_coefficients(res_bootstrap, labels_coefficients)

