


# Install packages if not already installed
install.packages(c("readxl", "dplyr", "janitor", "tidyverse", "ggplot2"))

##### Install packages if not already installed  ####
install.packages(c("readxl", "dplyr", "janitor", "tidyverse"))

# Load libraries
library(readxl)    # Excel File read
library(dplyr)     # Data manipulation
library(janitor)   # Frequency tables, clean tables
library(tidyverse) # Data handling + ggplot (visualization)


# Load libraries
library(readxl)    # Excel import
library(dplyr)     # Data manipulation
library(janitor)   # Frequency tables & % formatting
library(tidyverse) # General data handling
library(ggplot2)   # Visualization



# Load libraries
library(readxl)
library(dplyr)
library(gtsummary)
library(gt)



#### dataset read ####

# File path ( path)
file_path <- "D:/paid task/The Digital Dimension of Gender-Based Violence in Bangladesh_ Challenges and Ways Forward .xlsx"

#  (default = first sheet)
data <- read_excel(file_path)

# Data structure 
str(data)
head(data)







#### Demographic Summary####

# Gender #
# Count & %
tabyl(data, `Q3. Gender`) %>% adorn_pct_formatting() %>% adorn_totals("row")


# Age Group #

tabyl(data, `Q2. Age Group`) %>% adorn_pct_formatting() %>% adorn_totals("row")


# Educational Qualification#

tabyl(data, `Q5. Educational Qualification`) %>% adorn_pct_formatting() %>% adorn_totals("row")










library(flextable)
library(officer)




# Load libraries
library(readxl)
library(dplyr)
library(gtsummary)
library(gt)
library(janitor)  # optional, clean names


# Load libraries
library(readxl)
library(dplyr)
library(gtsummary)
library(gt)
library(janitor)
library(flextable)
library(officer)

# Clean column names
data <- data %>% clean_names()  # converts spaces/dots → underscores, lowercase

# Combined Demographic Summary Table
demo_table <- data %>%
  select(q3_gender, q2_age_group, q5_educational_qualification) %>%
  tbl_summary(
    by = NULL,  # overall summary
    statistic = all_categorical() ~ "{n} ({p}%)",
    missing_text = "0"
  )

# Convert to flextable for Word
demo_flex <- as_flex_table(demo_table) %>%
  set_caption("Demographic Summary")

# Optional formatting
demo_flex <- demo_flex %>%
  autofit() %>%
  bold(i = 1, bold = TRUE)  # bold header row

# Export to Word
read_docx() %>%
  body_add_flextable(demo_flex) %>%
  print(target = "Demographic_Summary.docx")







library(janitor)
data <- data %>% clean_names()
names(data)


#### DGBV Frequency####


library(janitor)
library(dplyr)

# -----------------------------
# 1️⃣ DGBV Frequency (Q9)
# -----------------------------
tabyl(data, q9_have_you_or_someone_you_know_experienced_digital_harassment_violence) %>%
  adorn_pct_formatting() %>%
  adorn_totals("row")

# -----------------------------
# 2️⃣ Likert-scale Summary (Q15)
# -----------------------------
# If factor
table(data$q15_do_you_think_social_media_increases_the_rate_of_online_gender_based_violence)

# Or, for nicer % output using janitor
tabyl(data, q15_do_you_think_social_media_increases_the_rate_of_online_gender_based_violence) %>%
  adorn_pct_formatting() %>%
  adorn_totals("row")









# -----------------------------
# Load libraries
# -----------------------------
library(readxl)
library(dplyr)
library(janitor)
library(flextable)
library(officer)

# -----------------------------
# Clean column names
# -----------------------------
data <- data %>% clean_names()

# -----------------------------
# 1️⃣ DGBV Frequency Table (Q9)
# -----------------------------
dgbv_freq <- tabyl(data, q9_have_you_or_someone_you_know_experienced_digital_harassment_violence) %>%
  adorn_pct_formatting() %>%
  adorn_totals("row")

# Convert to flextable
dgbv_flex <- flextable(dgbv_freq) %>%
  set_caption("Table 1. DGBV Experience Frequency (Q9)") %>%
  autofit() %>%
  bold(i = 1, bold = TRUE)

# -----------------------------
# 2️⃣ Likert-scale Summary Table (Q15)
# -----------------------------
# Count & % table
likert_freq <- tabyl(data, q15_do_you_think_social_media_increases_the_rate_of_online_gender_based_violence) %>%
  adorn_pct_formatting() %>%
  adorn_totals("row")

# Mean & Median
# Convert factor to numeric if not already
likert_levels <- c("Strongly Disagree" = 1,
                   "Disagree" = 2,
                   "Neutral" = 3,
                   "Agree" = 4,
                   "Strongly Agree" = 5)

data$likert_q15_num <- as.numeric(recode(data$q15_do_you_think_social_media_increases_the_rate_of_online_gender_based_violence, !!!likert_levels))

likert_mean <- round(mean(data$likert_q15_num, na.rm = TRUE), 2)
likert_median <- median(data$likert_q15_num, na.rm = TRUE)

# Convert to flextable
likert_flex <- flextable(likert_freq) %>%
  set_caption(paste0("Table 2. Perception: Social Media and Online GBV (Q15)\nMean = ", likert_mean, ", Median = ", likert_median)) %>%
  autofit() %>%
  bold(i = 1, bold = TRUE)

# -----------------------------
# 3️⃣ Export to MS Word
# -----------------------------
doc <- read_docx()

doc <- doc %>%
  body_add_par("DGBV Frequency and Likert Question Summary", style = "heading 2") %>%
  body_add_flextable(dgbv_flex) %>%
  body_add_par("", style = "Normal") %>%
  body_add_flextable(likert_flex)

# Save Word document
print(doc, target = "DGBV_Frequency_Likert_Summary.docx")















#### Crosstab ####

# -----------------------------
# Load libraries
# -----------------------------
library(dplyr)
library(janitor)
library(flextable)
library(officer)
library(vcd)  # for assocstats (Cramér's V)

# -----------------------------
# Ensure clean names
# -----------------------------
data <- data %>% clean_names()

# -----------------------------
# 1️⃣ Cross-tab: Gender × DGBV
# -----------------------------
gender_dgbv <- table(data$q3_gender, data$q9_have_you_or_someone_you_know_experienced_digital_harassment_violence)
gender_dgbv_pct <- round(prop.table(gender_dgbv, 2) * 100, 1)

# Fisher's test
gender_test <- fisher.test(gender_dgbv)

# Cramér's V (association strength)
gender_assoc <- assocstats(gender_dgbv)$cramer

# -----------------------------
# 2️⃣ Cross-tab: Urban vs Rural × DGBV
# -----------------------------
urban_dgbv <- table(data$q8_where_do_you_live, data$q9_have_you_or_someone_you_know_experienced_digital_harassment_violence)
urban_dgbv_pct <- round(prop.table(urban_dgbv, 2) * 100, 1)

# Fisher's test
urban_test <- fisher.test(urban_dgbv)

# Cramér's V
urban_assoc <- assocstats(urban_dgbv)$cramer

# -----------------------------
# 3️⃣ Sample size calculation (Cochran formula)
# -----------------------------
Z <- 1.96
p <- 0.5
e <- 0.05
n <- (Z^2 * p * (1 - p)) / e^2

# -----------------------------
# 4️⃣ Prepare Flextables
# -----------------------------
# Gender
gender_df <- as.data.frame.matrix(gender_dgbv)
gender_df_pct <- as.data.frame.matrix(gender_dgbv_pct)
gender_combined <- cbind(Gender = rownames(gender_df), gender_df, gender_df_pct)
colnames(gender_combined) <- c("Gender", paste0("Count: ", colnames(gender_df)), paste0("Pct: ", colnames(gender_df_pct)))
gender_flex <- flextable(gender_combined) %>%
  set_caption("Table 1. Gender × DGBV Experience (Counts & Column-wise %)") %>%
  autofit()

# Urban
urban_df <- as.data.frame.matrix(urban_dgbv)
urban_df_pct <- as.data.frame.matrix(urban_dgbv_pct)
urban_combined <- cbind(Origin = rownames(urban_df), urban_df, urban_df_pct)
colnames(urban_combined) <- c("Origin", paste0("Count: ", colnames(urban_df)), paste0("Pct: ", colnames(urban_df_pct)))
urban_flex <- flextable(urban_combined) %>%
  set_caption("Table 2. Urban vs Rural × DGBV Experience (Counts & Column-wise %)") %>%
  autofit()

# -----------------------------
# 5️⃣ Export to MS Word
# -----------------------------
doc <- read_docx() %>%
  body_add_par("Cross-tab Analysis: Gender & Urban Origin with DGBV", style = "heading 2") %>%
  body_add_par("Table 1: Gender × DGBV Experience", style = "heading 3") %>%
  body_add_flextable(gender_flex) %>%
  body_add_par("", style = "Normal") %>%
  body_add_par(paste0("Fisher's test p-value = ", signif(gender_test$p.value, 3), 
                      "; Cramér's V = ", round(gender_assoc, 2)), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  body_add_par("Table 2: Urban vs Rural × DGBV Experience", style = "heading 3") %>%
  body_add_flextable(urban_flex) %>%
  body_add_par("", style = "Normal") %>%
  body_add_par(paste0("Fisher's test p-value = ", signif(urban_test$p.value, 3), 
                      "; Cramér's V = ", round(urban_assoc, 2)), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  body_add_par(paste0("Sample size required (Cochran formula, 95% CI, 5% margin): n = ", round(n)), style = "Normal")

# Save Word document
print(doc, target = "DGBV_CrossTab_Fisher_Cramers.docx")












# -----------------------------
# Education × Reporting Behavior
# -----------------------------




names(data)







# -----------------------------
# Education × Reporting Behavior
# -----------------------------
# Cross-tab
edu_report <- table(data$q5_educational_qualification,
                    data$q12_have_you_ever_reported_any_dgbv_to_any_authority_police_cyber_unit_ngo_platform)

# Chi-square or Fisher test (auto)
if (any(edu_report < 5)) {
  edu_test <- fisher.test(edu_report)
  test_used <- "Fisher's Exact Test"
} else {
  edu_test <- chisq.test(edu_report)
  test_used <- "Chi-square Test"
}



















# =============================
# Education × Reporting Behavior Analysis
# =============================

# Cross-tab
edu_report <- table(
  data$q5_educational_qualification,
  data$q12_have_you_ever_reported_any_dgbv_to_any_authority_police_cyber_unit_ngo_platform
)

# Chi-square or Fisher test
if (any(edu_report < 5)) {
  edu_test <- fisher.test(edu_report)
  test_used <- "Fisher's Exact Test"
} else {
  edu_test <- chisq.test(edu_report)
  test_used <- "Chi-square Test"
}

cat("Test used:", test_used, "\n")
print(edu_test)







# =============================
# Education × Reporting Behavior Analysis
# =============================

# Cross-tab
edu_report <- table(
  data$q5_educational_qualification,
  data$q12_have_you_ever_reported_any_dgbv_to_any_authority_police_cyber_unit_ngo_platform
)

# Chi-square or Fisher test
if (any(edu_report < 5)) {
  edu_test <- fisher.test(edu_report)
  test_used <- "Fisher's Exact Test"
} else {
  edu_test <- chisq.test(edu_report)
  test_used <- "Chi-square Test"
}

cat("Test used:", test_used, "\n")
print(edu_test)








library(flextable)
library(officer)
library(dplyr)

# -----------------------------
# Gender Flextable
# -----------------------------
gender_df <- as.data.frame.matrix(gender_dgbv)
gender_df_pct <- as.data.frame.matrix(gender_dgbv_pct)
gender_combined <- cbind(Gender = rownames(gender_df), gender_df, gender_df_pct)
colnames(gender_combined) <- c("Gender", paste0("Count: ", colnames(gender_df)), paste0("Pct: ", colnames(gender_df_pct)))
gender_flex <- flextable(gender_combined) %>%
  set_caption("Table 1. Gender × DGBV Experience (Counts & Column-wise %)") %>%
  autofit() %>%
  bold(part = "header") %>%
  fontsize(size = 10, part = "all") %>%
  align(align = "center", part = "all")

# -----------------------------
# Urban Flextable
# -----------------------------
urban_df <- as.data.frame.matrix(urban_dgbv)
urban_df_pct <- as.data.frame.matrix(urban_dgbv_pct)
urban_combined <- cbind(Origin = rownames(urban_df), urban_df, urban_df_pct)
colnames(urban_combined) <- c("Origin", paste0("Count: ", colnames(urban_df)), paste0("Pct: ", colnames(urban_df_pct)))
urban_flex <- flextable(urban_combined) %>%
  set_caption("Table 2. Urban vs Rural × DGBV Experience (Counts & Column-wise %)") %>%
  autofit() %>%
  bold(part = "header") %>%
  fontsize(size = 10, part = "all") %>%
  align(align = "center", part = "all")

# -----------------------------
# Education × Reporting Flextable
# -----------------------------
edu_df <- as.data.frame.matrix(edu_report)
edu_combined <- cbind(Education = rownames(edu_df), edu_df)
edu_flex <- flextable(edu_combined) %>%
  set_caption("Table 3. Education × Reporting Behavior (Counts)") %>%
  autofit() %>%
  bold(part = "header") %>%
  fontsize(size = 10, part = "all") %>%
  align(align = "center", part = "all")

# -----------------------------
# Export all to Word
# -----------------------------
doc <- read_docx() %>%
  body_add_par("Cross-tab Analyses of DGBV", style = "heading 1") %>%
  
  # Gender table
  body_add_par("Table 1: Gender × DGBV Experience", style = "heading 2") %>%
  body_add_flextable(gender_flex) %>%
  body_add_par(paste0("Fisher's test p-value = ", signif(gender_test$p.value, 3), 
                      "; Cramér's V = ", round(gender_assoc, 2)), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  
  # Urban table
  body_add_par("Table 2: Urban vs Rural × DGBV Experience", style = "heading 2") %>%
  body_add_flextable(urban_flex) %>%
  body_add_par(paste0("Fisher's test p-value = ", signif(urban_test$p.value, 3), 
                      "; Cramér's V = ", round(urban_assoc, 2)), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  
  # Education table
  body_add_par("Table 3: Education × Reporting Behavior", style = "heading 2") %>%
  body_add_flextable(edu_flex) %>%
  body_add_par(paste0(test_used, " p-value = ", signif(edu_test$p.value, 3)), style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  
  # Optional: sample size note
  body_add_par(paste0("Sample size used: n = ", round(n)), style = "Normal")

# Save Word document
print(doc, target = "DGBV_CrossTab_AllTables.docx")







































#### Spearman Correlation Analysis ####




library(dplyr)
library(janitor)



# Function to safely convert Likert / ordinal to numeric
convert_likert <- function(x){
  x <- as.character(x)
  
  # Standard Likert mapping for agreement scale
  agree_levels <- c("Strongly Disagree", "Disagree", "Neutral", "Agree", "Strongly Agree")
  
  # Effectiveness scale
  effect_levels <- c("Not effective at all", "Less effective", "Neutral", "Effective", "Very Effective")
  
  # Yes/No/Don’t know
  yn_levels <- c("No", "Yes", "Don’t know", "Not adequate", "Adequate")
  
  # Choose which mapping to use
  if(all(x %in% agree_levels, na.rm = TRUE)) {
    as.numeric(factor(x, levels = agree_levels, ordered = TRUE))
  } else if(all(x %in% effect_levels, na.rm = TRUE)) {
    as.numeric(factor(x, levels = effect_levels, ordered = TRUE))
  } else if(all(x %in% yn_levels, na.rm = TRUE)) {
    as.numeric(factor(x, levels = yn_levels, ordered = TRUE))
  } else {
    # If mixed or unknown, try general factor
    as.numeric(factor(x))
  }
}

# Apply to each Likert variable
for(v in likert_vars){
  data[[v]] <- convert_likert(data[[v]])
}

# Check
sapply(data[likert_vars], function(x) table(x, useNA="ifany"))






























# Likert + demographic numeric variables
cor_vars <- c("age_num", "edu_num", likert_vars)

# Spearman correlation
cor_matrix <- cor(data[cor_vars], use = "pairwise.complete.obs", method = "spearman")

# P-values calculation
library(Hmisc)
cor_test <- rcorr(as.matrix(data[cor_vars]), type = "spearman")
cor_pval <- cor_test$P

# Round for display
cor_matrix <- round(cor_matrix, 3)
cor_pval <- round(cor_pval, 3)

# Print
cat("\nSpearman correlation matrix:\n")
print(cor_matrix)

cat("\nSpearman correlation p-value matrix:\n")
print(cor_pval)






# ===============================
# Required libraries
# ===============================
library(ggplot2)
library(reshape2)
library(dplyr)

# ===============================
# 1️⃣ Prepare correlation and p-value matrices
# ===============================
# numeric variables for correlation
num_vars <- c("age_num", "edu_num", 
              "Q15. Do you think social media increases the rate of online gender-based violence?",
              "Q11. How effective do you think current reporting mechanisms (police stations, online portals, hotlines) are in addressing DGBV?",
              "Q14. Do you think victims from marginalized groups (rural women, minorities) face more difficulty in reporting cases?",
              "Q17. Do you think social media platforms are helpful in preventing online GBV?",
              "Q18. Do you believe social media safety features (blocking, reporting, privacy settings) are effective and keep you safe?",
              "Q20. Do you think the existing laws and policies in Bangladesh are adequate to prevent DGBV?")

# correlation matrix
cor_mat <- cor(data[num_vars], use="pairwise.complete.obs", method="spearman")

# p-value matrix function
cor_pval <- function(df) {
  mat <- matrix(NA, ncol=ncol(df), nrow=ncol(df))
  colnames(mat) <- colnames(df)
  rownames(mat) <- colnames(df)
  for(i in 1:ncol(df)){
    for(j in 1:ncol(df)){
      x <- df[[i]]
      y <- df[[j]]
      if(all(is.na(x)) | all(is.na(y))){
        mat[i,j] <- NA
      } else {
        mat[i,j] <- cor.test(x, y, method="spearman")$p.value
      }
    }
  }
  return(mat)
}

p_mat <- cor_pval(data[num_vars])

# ===============================
# 2️⃣ Melt matrices for ggplot
# ===============================
cor_df <- melt(cor_mat, na.rm = TRUE)
p_df <- melt(p_mat, na.rm = TRUE)

# merge correlation and p-values
plot_df <- cor_df %>% 
  rename(Corr= value) %>% 
  left_join(p_df %>% rename(Pval=value), by=c("Var1","Var2"))

# ===============================
# 3️⃣ Plot heatmap
# ===============================
ggplot(plot_df, aes(Var2, Var1, fill=Corr)) +
  geom_tile(color="white") +
  scale_fill_gradient2(low="blue", high="red", mid="white", 
                       midpoint=0, limit=c(-1,1), space="Lab",
                       name="Spearman\nCorrelation") +
  geom_text(aes(label=ifelse(Pval<0.05, sprintf("%.2f*", Corr), sprintf("%.2f", Corr))),
            color="black", size=3) +
  theme_minimal() + 
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust=1)) +
  coord_fixed() +
  labs(x="", y="", title="Spearman Correlation Heatmap ( * = p < 0.05 )")












# 1️⃣ Load libraries
library(ggplot2)
library(reshape2)
library(dplyr)

# 2️⃣ Numeric variables
num_vars <- c("age_num", "edu_num", 
              "Q15. Do you think social media increases the rate of online gender-based violence?",
              "Q11. How effective do you think current reporting mechanisms (police stations, online portals, hotlines) are in addressing DGBV?",
              "Q14. Do you think victims from marginalized groups (rural women, minorities) face more difficulty in reporting cases?",
              "Q17. Do you think social media platforms are helpful in preventing online GBV?",
              "Q18. Do you believe social media safety features (blocking, reporting, privacy settings) are effective and keep you safe?",
              "Q20. Do you think the existing laws and policies in Bangladesh are adequate to prevent DGBV?")

# 3️⃣ Correlation matrix
cor_mat <- cor(data[num_vars], use="pairwise.complete.obs", method="spearman")

# 4️⃣ P-value matrix function
cor_pval <- function(df) {
  mat <- matrix(NA, ncol=ncol(df), nrow=ncol(df))
  colnames(mat) <- colnames(df)
  rownames(mat) <- colnames(df)
  for(i in 1:ncol(df)){
    for(j in 1:ncol(df)){
      x <- df[[i]]
      y <- df[[j]]
      if(all(is.na(x)) | all(is.na(y))){
        mat[i,j] <- NA
      } else {
        mat[i,j] <- cor.test(x, y, method="spearman")$p.value
      }
    }
  }
  return(mat)
}

p_mat <- cor_pval(data[num_vars])

# 5️⃣ Melt matrices
cor_df <- melt(cor_mat, na.rm = TRUE)
p_df <- melt(p_mat, na.rm = TRUE)

# Merge correlation and p-value
cor_df$pvalue <- p_df$value

# Add significance stars
cor_df$sig <- cut(cor_df$pvalue,
                  breaks=c(-Inf, 0.001, 0.01, 0.05, Inf),
                  labels=c("***", "**", "*", ""))

# 6️⃣ Heatmap plot
ggplot(cor_df, aes(x=Var1, y=Var2, fill=value)) +
  geom_tile(color="white") +
  scale_fill_gradient2(low="blue", high="red", mid="white",
                       midpoint=0, limit=c(-1,1), space="Lab",
                       name="Spearman\nCorrelation") +
  geom_text(aes(label=paste0(round(value,2), sig)), color="black", size=4) +
  theme_minimal() +
  theme(axis.text.x=element_text(angle=45, vjust=1, hjust=1)) +
  coord_fixed() +
  labs(x="", y="", title="Spearman Correlation Heatmap with p-values")



















# 1️⃣ Load libraries
library(ggplot2)
library(reshape2)
library(dplyr)

# 2️⃣ Numeric variables + short names
num_vars <- c("age_num", "edu_num", 
              "Q15. Do you think social media increases the rate of online gender-based violence?",
              "Q11. How effective do you think current reporting mechanisms (police stations, online portals, hotlines) are in addressing DGBV?",
              "Q14. Do you think victims from marginalized groups (rural women, minorities) face more difficulty in reporting cases?",
              "Q17. Do you think social media platforms are helpful in preventing online GBV?",
              "Q18. Do you believe social media safety features (blocking, reporting, privacy settings) are effective and keep you safe?",
              "Q20. Do you think the existing laws and policies in Bangladesh are adequate to prevent DGBV?")

short_names <- c("Age", "Edu", "SM_Violence", "Reporting_Eff", 
                 "Marginal_Diff", "SM_Prevent", "Safety_Features", "Law_Adequacy")

# 3️⃣ Correlation matrix
cor_mat <- cor(data[num_vars], use="pairwise.complete.obs", method="spearman")

# 4️⃣ Melt matrix for ggplot
cor_df <- melt(cor_mat, na.rm=TRUE)
cor_df$Var1 <- factor(cor_df$Var1, levels=num_vars, labels=short_names)
cor_df$Var2 <- factor(cor_df$Var2, levels=num_vars, labels=short_names)

# 5️⃣ Heatmap plot without value labels
ggplot(cor_df, aes(x=Var1, y=Var2, fill=value)) +
  geom_tile(color="white") +
  scale_fill_gradient2(low="blue", high="red", mid="white",
                       midpoint=0, limit=c(-1,1), space="Lab",
                       name="Spearman\nCorrelation") +
  theme_minimal() +
  theme(axis.text.x=element_text(angle=45, vjust=1, hjust=1, size=10),
        axis.text.y=element_text(size=10)) +
  coord_fixed() +
  labs(x="", y="", title="Spearman Correlation Heatmap (short names)")




# ===============================
# 1️⃣ Correlation matrix (numeric variables)
# ===============================
cor_mat <- cor(data[num_vars], use="pairwise.complete.obs", method="spearman")

# Round for table
cor_table <- round(cor_mat, 2)

# Convert to tibble for nicer printing
library(tibble)
cor_table_tibble <- as_tibble(cor_table, rownames = "Variable")
cor_table_tibble





library(ggplot2)
library(reshape2)
library(dplyr)

# Numeric vars & short names
num_vars <- c("age_num", "edu_num", 
              "Q15. Do you think social media increases the rate of online gender-based violence?",
              "Q11. How effective do you think current reporting mechanisms (police stations, online portals, hotlines) are in addressing DGBV?",
              "Q14. Do you think victims from marginalized groups (rural women, minorities) face more difficulty in reporting cases?",
              "Q17. Do you think social media platforms are helpful in preventing online GBV?",
              "Q18. Do you believe social media safety features (blocking, reporting, privacy settings) are effective and keep you safe?",
              "Q20. Do you think the existing laws and policies in Bangladesh are adequate to prevent DGBV?")

short_names <- c("Age", "Edu", "SM_Violence", "Reporting_Eff", 
                 "Marginal_Diff", "SM_Prevent", "Safety_Features", "Law_Adequacy")

# Correlation
cor_mat <- cor(data[num_vars], use="pairwise.complete.obs", method="spearman")

# Melt full matrix
cor_df <- melt(cor_mat, na.rm=TRUE)
cor_df$Var1 <- factor(cor_df$Var1, levels=num_vars, labels=short_names)
cor_df$Var2 <- factor(cor_df$Var2, levels=num_vars, labels=short_names)

# Full heatmap
ggplot(cor_df, aes(x=Var1, y=Var2, fill=value)) +
  geom_tile(color="white", size=0.3) +
  scale_fill_gradient2(low="#377eb8", mid="white", high="#e41a1c",
                       midpoint=0, limit=c(-1,1), space="Lab",
                       name="Spearman\nCorrelation") +
  theme_minimal(base_size = 14) +
  theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1, size=12, face="bold"),
        axis.text.y = element_text(size=12, face="bold"),
        panel.grid = element_blank(),
        legend.title = element_text(size=12, face="bold"),
        legend.text = element_text(size=10)) +
  coord_fixed() +
  labs(title="Spearman Correlation Heatmap",
       x="", y="")










library(flextable)
library(dplyr)

# Round correlation matrix to 2 decimals for table
cor_table_rounded <- round(cor_table, 2)

# Convert to tibble with variable names as row
cor_table_tibble <- as_tibble(cor_table_rounded, rownames = "Variable")

# Create flextable
cor_ft <- flextable(cor_table_tibble) %>%
  autofit() %>%                     # Auto-adjust column width
  theme_vanilla() %>%               # Clean minimal theme
  bold(part = "header") %>%         # Bold header
  set_header_labels(
    Variable = "Variable",
    age_num = "Age",
    edu_num = "Education",
    `Q15. Do you think social media increases the rate of online gender-based violence?` = "Social Media GBV ↑",
    `Q11. How effective do you think current reporting mechanisms (police stations, online portals, hotlines) are in addressing DGBV?` = "Reporting Effectiveness",
    `Q14. Do you think victims from marginalized groups (rural women, minorities) face more difficulty in reporting cases?` = "Marginalized Victims Difficulty",
    `Q17. Do you think social media platforms are helpful in preventing online GBV?` = "Social Media Preventive",
    `Q18. Do you believe social media safety features (blocking, reporting, privacy settings) are effective and keep you safe?` = "Safety Features Effective",
    `Q20. Do you think the existing laws and policies in Bangladesh are adequate to prevent DGBV?` = "Laws Adequate"
  ) %>%
  color(part = "body", color = "black") %>%  # Body text color
  fontsize(size = 10, part = "all") %>%     # Font size
  align(align = "center", part = "all")     # Center align

cor_ft



library(officer)
library(flextable)

# Create a new Word document
doc <- read_docx()

# Add a title or heading
doc <- doc %>%
  body_add_par("Table X. Spearman Correlation Matrix of Study Variables", style = "heading 2") %>%
  body_add_par("Note. Values are Spearman's rho. Variable names are abbreviated for clarity.", style = "Normal")

# Add the flextable
doc <- doc %>%
  body_add_flextable(cor_ft)

# Save Word document
print(doc, target = "Correlation_Table.docx")















