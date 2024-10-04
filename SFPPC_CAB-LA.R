#################################################
#                                               #  
#               CAB-LA ANALYSES                 #
#                                               #
#                  MJ Heise                     #
#                June 4, 2024                   #
#                                               #
#################################################

# CODE DESCRIPTION: This code reads in cabotegravir long-acting injection data 
# to examine injection schedule adherence in clinics in San Francisco, CA.
#
# 1. DATA CLEANING:
# Create new variables for on-time CAB injections and duration on CAB.
#
# 2. DEMOGRAPHICS:
# Age, race, ethnicity, and gender of participants. 
# 
# 3. RETENTION ON CAB:
# Fit survival curve to examine retention on CAB-LA; examine retention via
# Cox regression.
#
# 4. LATE INJECTIONS:
# Examine predictors of late injections in logistic mixed effects model. 


# Libraries
library(tidyverse) # v. 2.0.0, data cleaning
library(readxl) # v. 1.4.3, read in excel files
library(table1) # v. 1.4.3, table1
library(survival) # v. 3.5-5, survival analysis
library(ggsurvfit) # v. 0.3.0, survival analysis
library(gtsummary) # v. 1.7.2, tbl_regression function
library(ggplot2) # v. 3.4.3, data visualization
library(ggplotify) # v. 0.1.2, powerpoint ggplot
library(officer) # V.0.3.15, powerpoint ggplot
library(rvg) # v. 0.2.5, powerpoint ggplot
library(lme4) # v. 1.1-34, linear mixed effects models
library(lmerTest) # v. 3.1-3, linear mixed effects models
library(emmeans) # v. 1.8.9, marginal means comparison
library(broom.mixed) # v. 0.2.9.5, tidy glmer model for plotting


# Function to save ggplot object in powerpoint slide
# -Input: The ggplot object that you want to save in a powerpoint.
# -Optional inputs: Specified width and height of the outputted graph.
#  If no arguments are specified, the graph will encompass the entire
#  powerpoint slide.
# -Notes: After running the function, a window will open that 
#  allows you to select the powerpoint file. The graph will
#  save on a new slide at the end of the powerpoint.

create_pptx <- function(plt = last_plot(), path = file.choose(), width = 0, height = 0){
  if(!file.exists(path)) {
    out <- read_pptx()
  } else {
    out <- read_pptx(path)
  }
  
  if (width != 0 & height != 0) {
    out %>%
      add_slide(layout = "Title and Content", master = "Office Theme") %>%
      ph_with(value = dml(ggobj = plt), location = ph_location(left = 0, top = 0,
                                                               width = width, height = height)) %>%
      print(target = path)
  } else {
    out %>%
      add_slide(layout = "Title and Content", master = "Office Theme") %>%
      ph_with(value = dml(ggobj = plt), location = ph_location_fullsize()) %>%
      print(target = path)
    
  }
  
}

# Function to calculate percentages and create bar plots with labels over bars and no y-axis ticks
create_percent_barplot <- function(data, variable, fill_color, title, x_label) {
  data %>%
    group_by(!!sym(variable)) %>%
    summarise(count = n()) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ggplot(aes_string(x = variable, y = "percentage")) +
    geom_bar(stat = "identity", fill = fill_color, color = "black") +
    geom_text(aes_string(label = "paste0(round(percentage, 1), '%')"), 
              vjust = -0.5, size = 3.5) +  # Add percentage labels above bars
    theme_minimal() +
    labs(title = title, x = x_label, y = "Percentage (%)") +
    theme(axis.text.y = element_blank(),    # Remove y-axis text (percentage values)
          axis.ticks.y = element_blank(),   # Remove y-axis ticks
          axis.title.y = element_blank())   # Remove y-axis title
}


#### 1. DATA CLEANING ####
# Read in data
datPrep <- read_excel('DPH-Wide CAB-LA PrEP Pts (through 05.17.24).xlsx')
datChart <- read_excel('W86CAB_ChartReview_20240521.xlsx', na = 'NA')

# Chart review: subset columns
datChart %>%
  filter(!is.na(MRN)) %>%
  mutate(last_date_chart = ymd(last_date)) %>%
  rename(insurance = Insurance,
         discontinuationReason = `Reason for discontinuation`,
         oralPrepHistory = `History of oral PrEP use`,
         oralPrepAtInitiation = `Current oral PrEP use at time of initiation`) %>%
  mutate(discontinued = case_when(`Discontinued CAB-LA?` == 'no' ~ 0,
                                  `Discontinued CAB-LA?` == 'yes' ~ 1,
                                  `Discontinued CAB-LA?` == 'transfer' ~ 0,
                                  is.na(`Discontinued CAB-LA?`) ~ 0,
                                  .default = NA),
         opiateUse = case_when(`Opiate use` == 'yes' ~ 1,
                               `Opiate use` == 'no' ~ 0,
                               .default = NA),
         alcoholUseDis = case_when(`Alcohol use disorder (diagnosed)` == 'yes' ~ 1,
                                   `Alcohol use disorder (diagnosed)` == 'no' ~ 0,
                                   .default = NA),
         stimUse = case_when(`Stimulant use` == 'yes' ~ 1,
                             `Stimulant use` == 'no' ~ 0,
                             .default = NA),
         mentalHealthDx = case_when(`Co-occurring psychiatric diagnosis` == 'yes' ~ 1,
                                    `Co-occurring psychiatric diagnosis` == 'no' ~ 0,
                                    .default = NA),
         housing = case_when(Housing == 'homeless' ~ 'unstable',
                             Housing == 'stable' ~ 'stable',
                             Housing == 'unstable' ~ 'unstable',
                             .default = NA),
         oralPrep = case_when(oralPrepHistory == 'no' ~ 'no prior PrEP',
                              oralPrepHistory == 'yes' & oralPrepAtInitiation == 'no' ~ 'prior PrEP, not current',
                              oralPrepAtInitiation == 'yes' ~ 'oral PrEP lead-in',
                              .default= NA)) %>%
  select(MRN, discontinued, last_date_chart, exclude,
         housing, insurance, discontinuationReason, oralPrepHistory, oralPrepAtInitiation, oralPrep,
         opiateUse, alcoholUseDis, stimUse, mentalHealthDx, `Psychiatric Dx`) -> datChart

# PrEP data: rename columns
datPrep %>%
  rename(PatientName = `Patient Name`) -> datPrep

# Merge data
dat <- merge(datPrep, datChart, by = 'MRN', all = T)

# PrEP: filter out City Clinic patients
dat %>%
  filter(`Hospital Location` != 'SFDPH City Clinic') %>%
  filter(exclude != 1) -> dat

# Recode variables
dat %>%
  mutate(race = case_when(grepl(x=`Patient Race`, 'Black') ~ 'Black',
                          grepl(x=`Patient Race`, 'Asian') ~ 'Asian/PacIs',
                          grepl(x=`Patient Race`, 'American Indian or Alaska Native') ~ 'Other',
                          grepl(x=`Patient Race`, 'Native Hawaiian or Other Pacific Islander') ~ 'Asian/PacIs',
                          grepl(x=`Patient Race`, 'White') ~ 'White',
                          grepl(x=`Patient Race`, 'Other') ~ 'Other',
                          grepl(x=`Patient Race`, 'Decline to Answer') ~ 'DeclineAns',
                          .default = NA),
         race_short = case_when(grepl(x=`Patient Race`, 'Black') ~ 'Black',
                          grepl(x=`Patient Race`, 'Asian') ~ 'Other',
                          grepl(x=`Patient Race`, 'American Indian or Alaska Native') ~ 'Other',
                          grepl(x=`Patient Race`, 'Native Hawaiian or Other Pacific Islander') ~ 'Other',
                          grepl(x=`Patient Race`, 'White') ~ 'White',
                          grepl(x=`Patient Race`, 'Other') ~ 'Other',
                          grepl(x=`Patient Race`, 'Decline to Answer') ~ NA,
                          .default = NA),
         gender = case_when(`Gender Identity` == 'Female' ~ 'Female',
                            `Gender Identity` == 'Male' ~ 'Male',
                            `Gender Identity` == 'Non-Binary/Gender Queer' ~ 'Other',
                            `Gender Identity` == 'Other' ~ 'Other',
                            `Gender Identity` == 'Transgender Female' ~ 'TransgenderFemale',
                            `Gender Identity` == 'Transgender Male' ~ 'Other',
                            .default = NA),
         date = ymd(substr(`Administration Date & Time`, 1, 10)),
         substanceSum = (opiateUse + alcoholUseDis + stimUse),
         dx_anxiety = case_when(grepl(x = `Psychiatric Dx`, 'anxiety') ~ 'Anxiety',
                                .default = NA),
         dx_depression = case_when(grepl(x = `Psychiatric Dx`, 'depressi') ~ 'Depression',
                                   .default = NA),
         dx_bipolar = case_when(grepl(x = `Psychiatric Dx`, 'bipolar') ~ 'BipolarDisorder',
                                .default = NA),
         dx_ocd = case_when(grepl(x = `Psychiatric Dx`, 'OCD') ~ 'OCD',
                            .default = NA),
         dx_personality = case_when(grepl(x = `Psychiatric Dx`, 'borderline') ~ 'PersonalityDisorder',
                                    .default = NA),
         dx_schizoPsychotic = case_when(grepl(x = `Psychiatric Dx`, 'schizo') ~ 'Schizophrenia/Psychotic',
                                        grepl(x = `Psychiatric Dx`, 'psychotic') ~ 'Schizophrenia/Psychotic', 
                                        .default = NA),
         dx_ptsd = case_when(grepl(x = `Psychiatric Dx`, 'PTSD') ~ 'PTSD',
                             grepl(x = `Psychiatric Dx`, 'post-trauma') ~ 'PTSD',
                             .default = NA),
         site = case_when(`Hospital Location` == 'SFDPH ZSFG Building 5' ~ 'ZSFG',
                                `Hospital Location` == 'SFDPH ZSFG Building 80/90' ~ 'ZSFG',
                                `Hospital Location` == 'SFDPH 5th Bryant' ~ 'Other',
                                `Hospital Location` == 'SFDPH Castro Mission Health Center' ~ 'Castro Mission',
                                `Hospital Location` == 'SFDPH Cole St. Youth Clinic' ~ 'Other',
                                `Hospital Location` == 'SFDPH Larkin Street Youth Clinic' ~ 'Other',
                                `Hospital Location` == 'SFDPH MXM Resource Center' ~ 'MXM Resource Center',
                                `Hospital Location` == 'SFDPH Tom Waddell Urban Health Center' ~ 'Other',
                                .default = NA),
         site = relevel(factor(site), ref = 'ZSFG'),
         zsfgother = case_when(site == 'ZSFG' ~ 'ZSFG',
                                 .default = 'Other'),
         ward86other = case_when(`Hospital Location` == 'SFDPH ZSFG Building 80/90' ~ 'ZSFG',
                                 .default = 'Other'),
         substanceUseAny = case_when(stimUse == 1 | opiateUse == 1 | alcoholUseDis == 1 ~ 1,
                                     .default = 0)) -> dat

# Add in CAB injection dates from Chart Review records, some injection dates 
# were not entered in Epic but were in Chart notes
dat %>%
  group_by(PatientName) %>%
  arrange(date) %>%
  mutate(first_date = min(date),
         last_date = max(date)) %>%
  mutate(last_date = case_when(last_date_chart > last_date ~ last_date_chart,
                                    .default = last_date)) -> dat

# Add in whether CAB injection was on-time or late
dat %>%
  group_by(PatientName) %>%
  arrange(date) %>%
  mutate(cab_daysSinceInj = as.numeric(str_sub(date - lag(date))),
         cab_lateInj = case_when(is.na(cab_daysSinceInj) ~ 0,
                                 as.numeric(cab_daysSinceInj) > 35 & row_number() == 2 ~ 1,
                                 as.numeric(cab_daysSinceInj) > 63 ~ 1,
                                 TRUE ~ 0)) -> dat


# Number of CAB injections received
dat %>%
  group_by(MRN) %>%
  summarise(nInjections = n()) -> temp

dat <- merge(dat, temp, by = 'MRN', all = T)

# For participants who had one injection, they had 35 days on PrEP, those with 
# 2 injections had 63 days on PrEP
# Save new dataset (datSurv) with one row per participant
dat %>%
  rename(censorStatus = discontinued) %>%
  mutate(days = case_when(nInjections == 1 ~ 35,
                          nInjections == 2 ~ 63,
                          .default = as.numeric(gsub(" days", "", last_date - first_date)))) %>%
  group_by(MRN) %>%
  slice(1) -> datSurv


#### 2. DEMOGRAPHICS ####
table1(~ `Age (Current)` + factor(`Gender Identity`) + factor(`Patient Ethnicity`) +
         factor(race) + factor(housing) + factor(stimUse) + factor(opiateUse) + 
         factor(mentalHealthDx) +  factor(substanceUseAny) + factor(substanceSum) +
         factor(oralPrep) + 
         factor(site) + factor(`Hospital Location`), data = datSurv)

# Mental health
data.frame(Diagnosis = c("Anxiety", "Depression", "Bipolar Disorder", "OCD", "Borderline Personality", "Schizophrenia/Psychotic", "PTSD", "Any Mental Health Dx"),
           Count = c(
             table(datSurv$dx_anxiety)[["Anxiety"]],
             table(datSurv$dx_depression)[["Depression"]],
             table(datSurv$dx_bipolar)[["BipolarDisorder"]],
             table(datSurv$dx_ocd)[["OCD"]],
             table(datSurv$dx_personality)[["PersonalityDisorder"]],
             table(datSurv$dx_schizoPsychotic)[["Schizophrenia/Psychotic"]],
             table(datSurv$dx_ptsd)[["PTSD"]],
             sum(table(datSurv$mentalHealthDx))))

# Number of on-time injections
table(dat$cab_lateInj)

# Range of injection data
range(dat$`Administration Date & Time`)

# Demographics bar plots
# Bar plot for oralPrep
p1 <- create_percent_barplot(datSurv, "oralPrep", "#b5e2ff", 
                             "Percentage of Oral PrEP Use", 
                             "Oral PrEP Use")

# Bar plot for mentalHealthDx
p2 <- create_percent_barplot(datSurv, "mentalHealthDx", "#b8d1ae", 
                             "Percentage of Mental Health Diagnosis", 
                             "Mental Health Diagnosis")

# Bar plot for substanceUseAny
p3 <- create_percent_barplot(datSurv, "substanceUseAny", "#ffd6b9", 
                             "Percentage of Any Substance Use", 
                             "Substance Use (Any)")

# Bar plot for housing
p4 <- create_percent_barplot(datSurv, "housing", "#dbbeff", 
                             "Percentage of housing instability", 
                             "Housing instability")
# Save demographics bar plots
create_pptx(p1)
create_pptx(p2)
create_pptx(p3)
create_pptx(p4)


#### 3. RETENTION ON CAB #### 
# Create survival object [Surv(dats, censorStatus) ~ 1] and fit kaplan-meier curve
km <- survfit(Surv(days, censorStatus) ~ 1, data = datSurv)

# Create km plot
kmPlot <- ggsurvfit(km) +
  labs(
    x = "Days",
    y = "Overall survival probability"
  ) + 
  add_confidence_interval(fill = "blue", color = NA) +  # Change confidence interval color to light blue
  add_risktable() +
  ylim(0, 1) +
  scale_x_continuous(limits = c(0, 550), breaks = c(0, 90, 180, 270, 360, 450, 540)) +
  theme_minimal() +
  theme(legend.position = "none")  # Hide the legend since we don't need it

# Convert the ggsurvfit plot to a ggplot object to ensure compatibility
kmPlot <- as.ggplot(kmPlot)

# Create a PowerPoint document
pptx_doc <- read_pptx()

# Add a slide
pptx_doc <- add_slide(pptx_doc, layout = "Title and Content", master = "Office Theme")

# Add the plot to the slide using the rvg package
pptx_doc <- ph_with(pptx_doc, dml(ggobj = kmPlot), location = ph_location_fullsize())

# Save the PowerPoint document
print(pptx_doc, target = "KaplanMeierPlot.pptx")

# Print the median survival time from raw data
datSurv %>%
  filter(censorStatus == 0) %>%
  ungroup() %>%
  summarise(median = median(days),
            se = std.error(days),
            low95 = median - 1.96*se,
            up95 = median + 1.96*se)

# Participants who discontinued (n = 19)
datSurv %>%
  group_by(censorStatus) %>%
  summarise(n = n())

# Predictors of CAB discontinuation (Cox regression)
# Set reference level
datSurv %>%
  mutate(race_refW = relevel(factor(race), ref = 'White'),
         gender_refM = relevel(factor(gender), ref = 'Male')) -> datSurv

fit1_dis <- coxph(Surv(days, censorStatus) ~ `Age (Current)` + gender_refM + race_refW +
                housing + substanceSum + mentalHealthDx + oralPrep + site, data = datSurv)

# Formatted table for predictors of late CAB injections
fit1_dis %>%
  tbl_regression(exp = T) 

# Figure for OR
# Extract model summary
model_summary <- tidy(fit1_dis, conf.int = TRUE, effects = "fixed") # Use broom.mixed

# Filter out the intercept (if you don't want it in the plot)
model_summary <- model_summary %>%
  filter(term != "(Intercept)")

# Calculate odds ratios and confidence intervals
model_summary <- model_summary %>%
  mutate(estimate = exp(estimate),  # Convert log-odds to odds ratios
         conf.low = exp(conf.low),  # Lower CI
         conf.high = exp(conf.high), # Upper CI
         yAxis = rev(1:n())) %>%  # Reverse order for plotting
  mutate(name = case_when(term == '`Age (Current)`' ~ 'Age',
                          term == 'gender_refMFemale' ~ 'Cisgender female (v. Male)',
                          term == 'gender_refMOther' ~ 'Other gender (v. Male)',
                          term == 'gender_refMTransgenderFemale' ~ 'Transgender female (v. Male)',
                          term == 'race_refWAsian/PacIs' ~ 'Asian (v. White)',
                          term == 'race_refWBlack' ~ 'Black (v. White)',
                          term == 'race_refWOther' ~ 'Other race (v. White)',
                          term == 'housingunstable' ~ 'Housing unstable (v. Stable)',
                          term == 'substanceSum' ~ 'Substance use',
                          term == 'mentalHealthDx' ~ 'Mental health diagnosis',
                          term == 'oralPreporal PrEP lead-in' ~ 'Oral PrEP lead-in (v. no PrEP)',
                          term == 'oralPrepprior PrEP, not current' ~ 'Prior but not current PrEP (v. no PrEP)',
                          term == 'siteCastro Mission' ~ 'Castro Mission (v. SFG)',
                          term == 'siteMXM Resource Center' ~ 'MXM (v. SFG)',
                          term == 'siteOther' ~ 'Other clinic (v. SFG)'),
         yAxis = length(model_summary$term) + 1-row_number())

# Define the desired order for the y-axis
desired_order <- c("Age",
                   "Cisgender female (v. Male)",
                   "Transgender female (v. Male)",
                   "Other gender (v. Male)",
                   "Asian (v. White)",
                   "Black (v. White)",
                   "Other race (v. White)",
                   "Housing unstable (v. Stable)",
                   "Substance use",
                   "Mental health diagnosis",
                   "Prior but not current PrEP (v. no PrEP)",
                   "Oral PrEP lead-in (v. no PrEP)",
                   "Castro Mission (v. SFG)",
                   "MXM (v. SFG)",
                   "Other clinic (v. SFG)")

# Set the factor levels for 'name' based on the desired order
model_summary <- model_summary %>%
  mutate(name = factor(name, levels = desired_order))

# Filter to show fewer variables
model_summary %>%
  filter(name == 'Age' | name == 'Housing unstable (v. Stable)' | name == 'Substance use' |
           name == 'Mental health diagnosis' | name == 'Oral PrEP lead-in (v. no PrEP)' |
           name == 'Prior but not current PrEP (v. no PrEP)') -> model_summary_short

# Plot the OR and CI using ggplot2
fit1_dis_plot <- ggplot(model_summary_short, aes(x = estimate, y = reorder(name, yAxis))) +
  geom_vline(aes(xintercept = 1), size = 0.25, linetype = "dashed") + # Reference line at OR = 1
  geom_errorbarh(aes(xmax = conf.high, xmin = conf.low), size = 0.5, height = 0.2, color = "gray50") + # Error bars
  geom_point(size = 3.5, color = "orange") + # Points for OR
  geom_text(aes(label = sprintf("%.2f", estimate)), hjust = -0.2, size = 3.5) + # Add OR labels next to points
  theme_bw() + # Clean theme
  theme(panel.grid.minor = element_blank()) + # Remove minor grid lines
  scale_x_continuous(breaks = seq(0.5, 2, 0.5)) + # Customize x-axis breaks
  coord_trans(x = "log10") + # Log scale for x-axis
  ylab("") + # Empty y-axis label
  xlab("Odds Ratio (log scale)") + # X-axis label
  ggtitle("Odds Ratios and Confidence Intervals from Model")

# Convert the ggsurvfit plot to a ggplot object to ensure compatibility
fit1_dis_plot <- as.ggplot(fit1_dis_plot)

# Save
create_pptx(fit1_dis__plot)


#### 4. LATE INJECTIONS ####
# Predictors of late CAB injections
# Set reference level
dat %>%
  mutate(race_refW = relevel(factor(race), ref = 'White'),
         gender_refM = relevel(factor(gender), ref = 'Male')) -> dat

# Fit glmer model for late cabotegravir injections
fit2_late <- glmer(cab_lateInj ~ `Age (At Administration)` + gender_refM + race_refW + 
                housing + substanceSum + mentalHealthDx + oralPrep + (1|MRN),
              family = binomial(link = "logit"),
              data = dat,
              glmerControl(optimizer = "bobyqa"))

summary(fit2_late)

# Formatted table for predictors of late CAB injections
fit2_late %>%
  tbl_regression(exp = T) 

# Figure for OR
# Extract model summary
model_summary <- tidy(fit2_late, conf.int = TRUE, effects = "fixed") # Use broom.mixed

# Filter out the intercept (if you don't want it in the plot)
model_summary <- model_summary %>%
  filter(term != "(Intercept)")

# Calculate odds ratios and confidence intervals
model_summary <- model_summary %>%
  mutate(estimate = exp(estimate),  # Convert log-odds to odds ratios
         conf.low = exp(conf.low),  # Lower CI
         conf.high = exp(conf.high), # Upper CI
         yAxis = rev(1:n())) %>%  # Reverse order for plotting
  mutate(name = case_when(term == '`Age (At Administration)`' ~ 'Age',
                          term == 'gender_refMFemale' ~ 'Cisgender female (v. Male)',
                          term == 'gender_refMOther' ~ 'Other gender (v. Male)',
                          term == 'gender_refMTransgenderFemale' ~ 'Transgender female (v. Male)',
                          term == 'race_refWAsian/PacIs' ~ 'Asian (v. White)',
                          term == 'race_refWBlack' ~ 'Black (v. White)',
                          term == 'race_refWOther' ~ 'Other race (v. White)',
                          term == 'housingunstable' ~ 'Housing unstable (v. Stable)',
                          term == 'substanceSum' ~ 'Substance use',
                          term == 'mentalHealthDx' ~ 'Mental health diagnosis',
                          term == 'oralPreporal PrEP lead-in' ~ 'Oral PrEP lead-in (v. no PrEP)',
                          term == 'oralPrepprior PrEP, not current' ~ 'Prior but not current PrEP (v. no PrEP)'),
         yAxis = 13-row_number())

# Define the desired order for the y-axis
desired_order <- c("Age",
                   "Cisgender female (v. Male)",
                   "Transgender female (v. Male)",
                   "Other gender (v. Male)",
                   "Asian (v. White)",
                   "Black (v. White)",
                   "Other race (v. White)",
                   "Housing unstable (v. Stable)",
                   "Substance use",
                   "Mental health diagnosis",
                   "Prior but not current PrEP (v. no PrEP)",
                   "Oral PrEP lead-in (v. no PrEP)")

# Set the factor levels for 'name' based on the desired order
model_summary <- model_summary %>%
  mutate(name = factor(name, levels = desired_order))

# Filtered variables for plot
model_summary %>%
  filter(name == 'Age' | name == 'Substance use' | name == 'Mental health diagnosis' |
           name == 'Prior but not current PrEP (v. no PrEP)' |
           name == 'Oral PrEP lead-in (v. no PrEP)' | name == 'Housing unstable (v. Stable)') -> model_summary_short

# Plot the OR and CI using ggplot2
fit2_late_plot <- ggplot(model_summary_short, aes(x = estimate, y = reorder(name, yAxis))) +
  geom_vline(aes(xintercept = 1), size = 0.25, linetype = "dashed") + # Reference line at OR = 1
  geom_errorbarh(aes(xmax = conf.high, xmin = conf.low), size = 0.5, height = 0.2, color = "gray50") + # Error bars
  geom_point(size = 3.5, color = "orange") + # Points for OR
  geom_text(aes(label = sprintf("%.2f", estimate)), hjust = -0.2, size = 3.5) + # Add OR labels next to points
  theme_bw() + # Clean theme
  theme(panel.grid.minor = element_blank()) + # Remove minor grid lines
  scale_x_continuous(breaks = seq(0.5, 2, 0.5)) + # Customize x-axis breaks
  coord_trans(x = "log10") + # Log scale for x-axis
  ylab("") + # Empty y-axis label
  xlab("Odds Ratio (log scale)") + # X-axis label
  ggtitle("Odds Ratios and Confidence Intervals from Model")

# Convert the ggsurvfit plot to a ggplot object to ensure compatibility
fit2_late_plot <- as.ggplot(fit2_late_plot)

# Save
create_pptx(fit2_late_plot)
