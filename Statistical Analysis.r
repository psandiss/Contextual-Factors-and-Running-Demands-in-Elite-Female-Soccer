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
df <- read_excel("your_file.xlsx", sheet = "your_sheet")
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
    PointsDifference < 3 ~ "Small",
    PointsDifference >= 3 & PointsDifference < 8 ~ "Moderate",
    PointsDifference >= 8 ~ "Large"
  ))

##############################################
# Descriptive statistics function
##############################################
descriptive_stats <- function(variable) {
  fullMatch %>%
    mutate(
      Position = factor(Position, 
                        levels = c("Defender", "Fullback", "Midfielder", "Forward")),
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
# Cross tables for contextual variables
##############################################
tables <- list()
for (var in independentVars) {
  for (depVar in dependentVars) {
    table_name <- paste0("table_", var, "_", depVar)  
    tables[[table_name]] <- calculate_table(var, depVar)
  }
}

# Save results in Excel (each table in a different sheet)
wb <- createWorkbook()
lapply(seq_along(tables), function(i) {
  sheet_name <- names(tables)[i]
  sheet_name <- substr(sheet_name, nchar(sheet_name) - 30, nchar(sheet_name))
  addWorksheet(wb, sheet_name)  
  writeData(wb, sheet = sheet_name, tables[[i]])  
})
saveWorkbook(wb, "IC_Days.xlsx", overwrite = TRUE)

##############################################
# Generalized Estimation Equations (GEE)
##############################################
dependentVars = c("DistRel") # Example: set dependent variables

for (depVar in dependentVars) {
  indepVar <- "Result" # Example: set independent variable
  
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
    MeanAcc = mean(ACCAbove2, na.rm = TRUE),
    SDAcc = sd(ACCAbove2, na.rm = TRUE),
    MeanDec = mean(DCCAbove2, na.rm = TRUE),
    SDDec = sd(DCCAbove2, na.rm = TRUE)
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
