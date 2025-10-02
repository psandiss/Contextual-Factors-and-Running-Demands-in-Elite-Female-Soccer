##############################################
# Load packages
##############################################
library(readxl)
library(car)
library(dplyr)
library(lme4)
library(emmeans)
library(redres)
library(geepack)
library(readr)
library(tidyr)
library(openxlsx)
library(afex)
library(ggplot2)
library(broom)
library(multcomp)
library(lsmeans)
library(effectsize)
library(tidyverse)
library(ggpubr)
library(Factoshiny)
library(gridExtra)
library(simr)
library(ggstatsplot)

##############################################
# Load dataset
##############################################
df <- read_excel("path/to/your_file.xlsx", sheet = "sheet")
df$idNumeric <- as.numeric(as.factor(df$Athlete))   # create numeric athlete ID

##############################################
# Split by halves and full match
##############################################
fullMatch <- df %>%
  filter(MatchPeriod == "Full") %>%
  arrange(idNumeric)

firstHalf <- df %>%
  filter(MatchPeriod == "1stHalf") %>%
  arrange(idNumeric)

secondHalf <- df %>%
  filter(MatchPeriod == "2ndHalf") %>%
  arrange(idNumeric)

##############################################
# Contextual categorical variables
##############################################
fullMatch <- fullMatch %>%
  mutate(MatchRecovery = case_when(
    DaysBetweenMatches <= 4 ~ "≤4 days",
    DaysBetweenMatches >= 5 & DaysBetweenMatches <= 7 ~ "5–7 days",
    DaysBetweenMatches >= 8 ~ "≥8 days"
  ))

fullMatch <- fullMatch %>%
  mutate(PointsDiffCat = case_when(
    PointDif < 3 ~ "Small",
    PointDif >= 3 & PointDif < 8 ~ "Moderate",
    PointDif >= 8 ~ "Large"
  ))

##############################################
# Descriptive statistics function
##############################################
descriptive_stats <- function(variable) {
  fullMatch %>%
    mutate(
      Position = factor(Position, 
                        levels = c("defender", "fullback", "midfielder", "foward")),
      ReportCondition = factor(ReportCondition, 
                               levels = c("90+", "45+", "5+"))
    ) %>%
    group_by(ReportCondition) %>%
    summarise(
      Mean = mean(.data[[variable]], na.rm = TRUE),
      SD = sd(.data[[variable]], na.rm = TRUE),
      .groups = "drop"
    ) %>%
    arrange(ReportCondition) %>%
    write_csv(paste0("summary_", variable, ".csv"))
}

##############################################
# Generalized Estimation Equations (GEE)
##############################################

dependentVars = c("DistanceRel") # Example: set dependent variables

for (depVar in dependentVars) {
  indepVar <- "PointsDiffCat" # Example: set independent variable
  
  formula <- as.formula(paste(depVar, " ~ Group * ", indepVar))
  
  # Exclude NAs
  data_temp <- fullMatch[!is.na(fullMatch[[indepVar]]), ]
  
  gee_model <- geeglm(
    formula,
    id = idNumeric,
    family = gaussian,
    corstr = "independence",
    data = data_temp
  )
  
  print(anova(gee_model))
  
  # Post-hoc comparisons
  efGroup <- lsmeans(gee_model, "Group", adjust = "tukey")
  print(pairs(efGroup))
  
  efVar <- lsmeans(gee_model, indepVar, adjust = "tukey")
  print(pairs(efVar))
  
  formula_inter <- as.formula(paste("~ Group | ", indepVar))
  EF_inter <- lsmeans(gee_model, formula_inter, adjust = "tukey")
  print(pairs(EF_inter, by = "Group"))
  print(pairs(EF_inter, by = indepVar))
}

##############################################
# Absolute descriptive summaries
##############################################
summaryAbs <- fullMatch %>%
  group_by(Group) %>%
  summarise(
    MeanTotalDist = mean(Distance, na.rm = TRUE),
    SDTotalDist = sd(Distance, na.rm = TRUE),
    MeanZone1 = mean(Dist0to7, na.rm = TRUE),
    SDZone1 = sd(Dist0to7, na.rm = TRUE),
    MeanZone2 = mean(Dist7to13, na.rm = TRUE),
    SDZone2 = sd(Dist7to13, na.rm = TRUE),
    MeanZone3 = mean(Dist13to19, na.rm = TRUE),
    SDZone3 = sd(Dist13to19, na.rm = TRUE),
    MeanZone4 = mean(Dist19to23, na.rm = TRUE),
    SDZone4 = sd(Dist19to23, na.rm = TRUE),
    MeanZone5 = mean(DistAbove23, na.rm = TRUE),
    SDZone5 = sd(DistAbove23, na.rm = TRUE),
    MeanZone5Effort = mean(HighEfforts23, na.rm = TRUE),
    SDZone5Effort = sd(HighEfforts23, na.rm = TRUE),
    MeanPlayerLoad = mean(PlayerLoad, na.rm = TRUE),
    SDPlayerLoad = sd(PlayerLoad, na.rm = TRUE),
    MeanAcc = mean(ACCACIMA2, na.rm = TRUE),
    SDAcc = sd(ACCACIMA2, na.rm = TRUE),
    MeanDec = mean(DCCACIMA2, na.rm = TRUE),
    SDDec = sd(DCCACIMA2, na.rm = TRUE)
  )

##############################################
# Example plots
##############################################
ggplot(summaryAbs, aes(x = Group, y = MeanTotalDist, fill = Group)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.85), width = 0.6, color = "black") +
  geom_errorbar(aes(ymin = MeanTotalDist, ymax = MeanTotalDist + SDTotalDist),
                position = position_dodge(width = 0.8), width = 0.25) +
  scale_fill_manual(values = c("black", "gray60")) +
  theme_minimal()

# Boxplot function
create_boxplot <- function(y_var, y_label, title, show_legend = TRUE) {
  ggplot(fullMatch, aes(x = Position, y = !!sym(y_var), fill = Group)) +
    geom_boxplot() +
    labs(y = y_label, title = title) +
    scale_fill_manual(values = c("Starter" = "#000000", "Non-starter" = "gray60")) +
    theme_minimal() +
    guides(fill = if (show_legend) guide_legend(title = NULL) else "none")
}

