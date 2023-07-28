
# ============================= IMPORT LIBRARIES ===============================

library(readxl)
library(MASS) # load before dplyr to not have 'select' function hidden
library(tidyr)
library(dplyr)
library(ggplot2) # graphs
library(lmerTest) # mixed models
library(gtsummary) #tables
library(tableone) # tables
library(ipw) # weights
library(survey) # apply weights to population (table and SMD)
library(docstring)

# ============================= IMPORT DATA ====================================

# set working directory
setwd("D:/Documents_D/INSA/4A/stage/etude_CRP_Lea_Lethuillier")

# path to export the tables
output_path = 'output_baseline/'

final_df = read_excel('data/final_dataframe_baseline.xlsx') #18,445

# list of covariates and their type
covariates = c('Age', 'ESS', 'BMI_class', 'Smoking', 'Sex_Male','Metabolic.Diabetes',
               'CV.Left.Ventricular.Hypertrophy', 'CV.hypertension',
               'CV.Ischemic.Heart.Disease','CV.TIA.or.stroke', 'CV.Cardiac.failure',
               'Pulmonary...COPD', 'Other...Neurological.Disease', 
               'Other...Psychiatric.Disease', 'Other...Inflammatory.Disease')

quanti_var = c('Age', 'ESS')
factor_var = covariates[!covariates %in% quanti_var]
binary_var = factor_var[-which(factor_var == 'BMI_class')]

# put variables as factor
final_df <- final_df %>%
  mutate_at(c(factor_var, 'AHI_class', 'AHI_ordinal', 'Site', 'Patient.ID'),
            ~ as.factor(.)) %>%
  mutate(BMI_class = factor(BMI_class, levels = c("<25", "25-30", "30-35", ">35"))) %>%
  mutate(AHI_class = relevel(AHI_class, ref = "No OSA")) %>%
  mutate(AHI_ordinal = factor(AHI_class, order = TRUE,
                              levels = c('No OSA', 'Mild OSA', 'Moderate OSA',
                                         'Severe OSA')))


# =========================== SENSIBILITY ANALYSES =============================

# Function to plot SMD after that the population has been weighted
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

subgroup_analysis <- function(df1, compute_weights, var = NULL){
  #' Subgroup analysis
  #' 
  #' Create a subgroup analysis with the known covariates and model
  #' @param df1 A dataframe with the variables
  #' @param compute_weights a function to compute weights for the model
  #' @param var The name of the variable used to create the subgroups
  
  
  # list of covariates and their type
  covariates = c('Age', 'ESS', 'BMI_class', 'Smoking', 'Sex_Male',
                 'Metabolic.Diabetes', 'CV.Left.Ventricular.Hypertrophy',
                 'CV.hypertension','CV.Ischemic.Heart.Disease','CV.TIA.or.stroke',
                 'CV.Cardiac.failure', 'Pulmonary...COPD', 'Other...Neurological.Disease', 
                 'Other...Psychiatric.Disease', 'Other...Inflammatory.Disease')
  
  quanti_var = c('Age', 'ESS')
  factor_var = covariates[!covariates %in% quanti_var]
  binary_var = factor_var[-which(factor_var == 'BMI_class')]
  
  # if the variable on which is done the subgroup analysis is specified, remove it
  if(!is.null(var)){
    covariates = covariates[-which(covariates == var)]
    if(var %in% factor_var){
      factor_var = factor_var[-which(factor_var == var)]
      if(var %in% binary_var){
        binary_var = binary_var[-which(binary_var == var)]
      }
    }
  }
  
  # compute weights
  df1 = compute_weights(df1) #function to implement for each subgroup analysis
  
  weights_by_group = df1 %>%
    group_by(AHI_ordinal) %>%
    summarise_at(vars(ipw), list(median = median, mean = mean, min = min, max = max))
  print(weights_by_group)
  
  # apply weights
  weighted_data = svydesign(ids = ~1, data = df1, weights = ~ipw)
  unweighted_table <-
    CreateTableOne(vars = covariates, factorVars = factor_var,
                   strata = "AHI_class", data = df1, test = TRUE)
  weighted_table <-
    svyCreateTableOne(vars = c(covariates), factorVars = factor_var,
                      strata = "AHI_class", data = weighted_data, test = TRUE)
  
  print(plot_SMD(unweighted_table, weighted_table))  
  
  # create model
  formula = as.formula(paste('C.Reactive.Protein.mg.L.','~ AHI_class + ', 
                             paste(covariates, collapse = '+'),
                             '+ (1|Site) + (1|Visit.Year)'))
  
  model <- lmer(formula, data= df1, weights = ipw)
  
  summary_model = model %>% tbl_regression(
    show_single_row = all_of(binary_var),
    pvalue_fun = ~ style_pvalue(.x, digits = 2)
  )
  
  list = c('BMI_class', 'AHI_class')
  if(!is.null(var)){
    if(var %in% list){
      list = list[-which(list == var)]
    }
  }
  summary_model = add_global_p(summary_model, include = all_of(list),
                               keep = TRUE)
  return(summary_model)
}

# -------------------------------- Analysis sex --------------------------------

# update of the function to compute the weights
# remove the sex variable from the denominator

compute_weights_sex <- function(df1){
  df1 = as.data.frame(df1)
  
  weights <- ipwpoint(
    exposure = AHI_ordinal,
    family = "ordinal",
    link = 'logit',
    numerator = ~1, #stabilized
    denominator = ~  Age + ESS + BMI_class + Smoking  + 
      Metabolic.Diabetes + CV.Left.Ventricular.Hypertrophy  + 
      CV.Ischemic.Heart.Disease  + CV.hypertension + CV.TIA.or.stroke +
      CV.Cardiac.failure + Pulmonary...COPD + Other...Neurological.Disease +
      Other...Psychiatric.Disease + Other...Inflammatory.Disease, 
    data = df1,
    trunc = .01
  )
  
  # add weights values to the dataframe
  df1 <- df1 %>% 
    mutate(ipw = weights$ipw.weights)%>%
    mutate(ipw_trunc = weights$weights.trunc)
  
  return(df1)
}

# create dataframes
female_df = final_df %>% filter(Sex_Male == 0) # 5,409
male_df = final_df %>% filter(Sex_Male == 1) # 13,036

female = subgroup_analysis(female_df, compute_weights_sex, 'Sex_Male')
male = subgroup_analysis(male_df,compute_weights_sex, 'Sex_Male')

comparison_sex = tbl_merge(
  list(female, male),
  tab_spanner = c('Women', 'Men')
)

# ----------------------------- Analysis COPD  ---------------------------------

# update of the function to compute the weights
# remove the COPD variable from the denominator

compute_weights_COPD <- function(df1){
  df1 = as.data.frame(df1)
  
  weights <- ipwpoint(
    exposure = AHI_ordinal,
    family = "ordinal",
    link = 'logit',
    numerator = ~1, #stabilized
    denominator = ~  Age + ESS + BMI_class + Sex_Male + Smoking  + 
      Metabolic.Diabetes + CV.Left.Ventricular.Hypertrophy  + 
      CV.Ischemic.Heart.Disease  + CV.hypertension + CV.TIA.or.stroke +
      CV.Cardiac.failure + Other...Neurological.Disease +
      Other...Psychiatric.Disease + Other...Inflammatory.Disease, 
    data = df1,
    trunc = .01
  )
  
  # add weights values to the dataframe
  df1 <- df1 %>% 
    mutate(ipw = weights$ipw.weights)%>%
    mutate(ipw_trunc = weights$weights.trunc)
  
  return(df1)
}

# create dataframe only with COPD patients
no_COPD_df = final_df %>% filter(Pulmonary...COPD == 0) #17,262
COPD_df = final_df %>% filter(Pulmonary...COPD == 1) # 1,183

no_COPD = subgroup_analysis(no_COPD_df, compute_weights_COPD, 'Pulmonary...COPD')
COPD = subgroup_analysis(COPD_df, compute_weights_COPD, 'Pulmonary...COPD')

comparison_COPD = tbl_merge(
  list(no_COPD, COPD),
  tab_spanner = c('no COPD', 'COPD')
)


# ----------------------------- Analysis PG/PSG --------------------------------

compute_weights_PG_PSG <- function(df1){
  df1 = as.data.frame(df1)
  
  # not possible to enter formulas dynamically in this function...
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
    data = df1,
    trunc = .01
  )
  
  # add weights values to the dataframe
  df1 <- df1 %>% 
    mutate(ipw = weights$ipw.weights)%>%
    mutate(ipw_trunc = weights$weights.trunc)
  
  return(df1)
}

# load the database to retrieve the type of test for each patient (PG or PSG)
load("data/bd_esada_initial.RData")
data = data %>% 
  filter(visit == 0 & Patient.ID %in% final_df$Patient.ID) %>%
  select(Patient.ID, Test...Attended.PSG, Test...Unattended.PSG, Test...Polygraphy)

data = data %>% mutate_at(names(data), ~ as.factor(.))

# join type of test with the dataframe with each patient
final_df = final_df %>% left_join(data, by = 'Patient.ID')

# create one dataframe for patients with PG and one df for patients with PSG
PG_df = final_df %>% filter(Test...Polygraphy == 1) #7,339
PSG_df = final_df %>%
  filter(Test...Attended.PSG == 1 | Test...Unattended.PSG == 1) #11,033

PG = subgroup_analysis(PG_df, compute_weights_PG_PSG)
PSG = subgroup_analysis(PSG_df, compute_weights_PG_PSG)

comparison_method = tbl_merge(
  list(PG, PSG),
  tab_spanner = c('PG', 'PSG')
)

