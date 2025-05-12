# Input variables:
# model2export = the fitted GLMM model object (list) you want to export from
# ExportFolder = the path to the folder where you want the Word file to be saved (use forward slashes)
# FileName     = the desired name for the output Word document (without .docx extension) 
# 
# Example of calling to the function: 
# GLMMsummary2doc(model2export=glmm_4paper_subset2,
#                 ExportFolder="Harod/tables 4 paper/",
#                 FileName="glmm_4paper_subset2_outputTable")


GLMMsummary2docx <- function(model2export,ExportFolder,FileName)
{
# ====================================================
# This script extracts fixed effects from a GLMM model 
# (***been tested with glmmTMB only, if you use a different statistical model, test it first***),
# formats the estimates, standard errors, test statistics,
# and p-values, adds significance stars, and exports 
# the result as a styled Word table using flextable.
# Requires: broom.mixed, dplyr, flextable, officer
# 
# Input variables:
# model2export = the fitted GLMM model object (list) you want to export from
# ExportFolder = the path to the folder where you want the Word file to be saved (use forward slashes)
# FileName     = the desired name for the output Word document (without .docx extension) 
# 
# Example of calling to the function: 
# GLMMsummary2doc(model2export=glmm_4paper_subset2,
#                 ExportFolder="Harod/tables 4 paper/",
#                 FileName="glmm_4paper_subset2_outputTable")
# ====================================================

library(broom.mixed)  # for tidy() on mixed models
library(flextable)    # for creating Word tables
library(officer)      # for exporting Word documents
library(dplyr)  # for data manipulation: mutate(), case_when(), and piping (%>%)

# Extract model summary (fixed effects only)
summary_glmm_4paper <- tidy( 
  model2export,       # <- user input: your fitted GLMM model object
  effects = "fixed", # extract only fixed effects (not random effects)
  conf.int = TRUE    # also extract confidence intervals (optional, not used below)
)

# Select and rename relevant columns (adjust if needed)
model_table_glmm_4paper <- summary_glmm_4paper[, c("term", "estimate", "std.error", "statistic", "p.value")] # insert here the existing names
colnames(model_table_glmm_4paper) <- c("Variable", "Estimate", "Std.Error", "z value", "p value")            # insert here the required names

# Step 1: Format numerical values and add significance stars
model_table_glmm_4paper <- model_table_glmm_4paper %>%
  dplyr::mutate(p_raw = `p value`) %>%   # Save original numeric p values for later use in significance annotation
  dplyr::mutate(
    Estimate = round(Estimate, 4),       # round estimates to 4 digits
    `Std.Error` = round(`Std.Error`, 4), # round standard errors to 4 digits
    `z value` = round(`z value`, 2),     # round z-values to 2 digits
    `p value` = dplyr::case_when(                               # Format p values with different rules:
      p_raw < 2.22e-16 ~ "< 2.22e-16",                          # - Show "< 2.22e-16" for extremely small values
      p_raw < 0.001 ~ formatC(p_raw, format = "e", digits = 2), # - Use scientific notation for small values (< 0.001)
      TRUE ~ formatC(p_raw, format = "f", digits = 4)           # - Otherwise, show 4 decimal places
    )
  ) %>% 
  dplyr::mutate(
    Signif = dplyr::case_when(  # create a new column with significance stars
      p_raw < 0.001 ~ "***",
      p_raw < 0.01  ~ "**",
      p_raw < 0.05  ~ "*",
      p_raw < 0.1   ~ ".",
      TRUE          ~ ""         # empty string if not significant
    )
  ) %>% 
  dplyr::select(-p_raw) # Remove the temporary raw p-value column (optional)

# Step 2: Create the flextable
ft <- flextable(model_table_glmm_4paper) %>%
  autofit() %>%           # automatically adjust column widths
  theme_booktabs()        # apply a clean professional theme to the table

# Step 3: Export the table to a Word document
doc <- read_docx()                                          # create a new Word document
doc <- body_add_flextable(doc, ft)                          # insert the table into the document
print(doc, target = paste0(ExportFolder,FileName,".docx"))  # <- user input: path and file name for the output Word file
}
      