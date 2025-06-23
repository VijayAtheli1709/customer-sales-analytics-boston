# Loading necessary packages
library(readxl)
library(readr)
library(writexl)
library(moments)
library(ggplot2)
library(dplyr)

# Load the sales data from the correct sheet
sales_data <- read_excel("sales.xlsx", sheet = "sales-data")

# Load stores data
stores_data <- read_csv("stores.csv")

# Clean unnecessary columns in stores_data if present
stores_data <- stores_data %>% select(store, city, state, sqft, store.tier)

# Merge store information into sales data
sales_data <- left_join(sales_data, stores_data, by = "store")

# View merged data
head(sales_data)

# Load customers data
customers_data <- customers

# Part A: Data Cleaning
# Clean customer state column
customers_data$customer.state[customers_data$customer.state == "Mass."] <- "MA"
customers_data$customer.state[customers_data$customer.state == "Massachusetts"] <- "MA"
customers_data$customer.state[customers_data$customer.state == "Mass"] <- "MA"
customers_data$customer.state[customers_data$customer.state == "Massachusets"] <- "MA"
customers_data$customer.state[customers_data$customer.state == "Connecticut"] <- "CT"
customers_data$customer.state[customers_data$customer.state == "Conn."] <- "CT"

# Clean selection satisfaction column
customers_data$selection[customers_data$selection == "" | is.na(customers_data$selection)] <- NA

# Clean birthday.month column
customers_data$birthday.month[customers_data$birthday.month == "October"] <- "10"
customers_data$birthday.month[customers_data$birthday.month == "March"] <- "3"
customers_data$birthday.month[customers_data$birthday.month == "Mar"] <- "3"
customers_data$birthday.month[customers_data$birthday.month == "February"] <- "2"
customers_data$birthday.month[customers_data$birthday.month == "Apr."] <- "4"
customers_data$birthday.month[customers_data$birthday.month == "Feb."] <- "2"
customers_data$birthday.month[customers_data$birthday.month == "April"] <- "4"
customers_data$birthday.month[customers_data$birthday.month == "Nov."] <- "11"
customers_data$birthday.month[customers_data$birthday.month == "July"] <- "7"
customers_data$birthday.month[customers_data$birthday.month == "Oct"] <- "10"
customers_data$birthday.month[customers_data$birthday.month == 0] <- NA        
customers_data <- customers_data[!is.na(customers_data$birthday.month), ]

# Aggregate sales data by customer
summary_table <- sales_data %>%
  group_by(customer.id) %>%
  summarize(
    total_items_purchased = n(),
    avg_item_sale_price = mean(sale.amount, na.rm = TRUE)
  )

# Merge with cleaned customer data
customer_purchases <- left_join(summary_table, customers_data, by = "customer.id")
write.csv(customer_purchases, "Outputs/customer_purchases.csv", row.names = FALSE)
write_xlsx(customer_purchases, "Outputs/customer_purchases.xlsx")

# Summary statistics
mean_sale <- mean(sales_data$sale.amount, na.rm = TRUE)
median_sale <- median(sales_data$sale.amount, na.rm = TRUE)
sd_sale <- sd(sales_data$sale.amount, na.rm = TRUE)
skew_sale <- skewness(sales_data$sale.amount, na.rm = TRUE)

# Create Visuals folder
if(!dir.exists("Visuals")) dir.create("Visuals")

# Boxplot for all sale amounts
boxplot_all <- ggplot(sales_data, aes(y = sale.amount)) +
  geom_boxplot(fill = "skyblue", color = "black") +
  ggtitle("Boxplot of Sale Amounts for All Sales") +
  ylab("Sale Amount") +
  theme_minimal()
ggsave("Visuals/boxplot_all_sales.png", boxplot_all, width = 8, height = 5)

# Boxplot by product category
boxplot_category <- ggplot(sales_data, aes(x = category, y = sale.amount, fill = category)) +
  geom_boxplot() +
  ggtitle("Boxplot of Sale Amounts by Product Category") +
  ylab("Sale Amount") +
  xlab("Product Category") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
ggsave("Visuals/boxplot_by_category.png", boxplot_category, width = 8, height = 5)

# Blended Gross Margin by Product Category
gross_margin_by_category <- sales_data %>%
  group_by(category) %>%
  summarise(
    total_sale_amount = sum(sale.amount, na.rm = TRUE),
    total_ext_cost = sum(ext.cost, na.rm = TRUE),
    blended_gross_margin = (total_sale_amount - total_ext_cost) / total_sale_amount * 100
  )
write.csv(gross_margin_by_category, "Outputs/gross_margin_by_category.csv", row.names = FALSE)

# Bar chart for gross margin by category
barplot_gm_category <- ggplot(gross_margin_by_category, aes(x = category, y = blended_gross_margin, fill = category)) +
  geom_bar(stat = "identity") +
  ggtitle("Blended Gross Margin by Product Category") +
  xlab("Product Category") +
  ylab("Blended Gross Margin (%)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
ggsave("Visuals/barplot_gross_margin_by_category.png", barplot_gm_category, width = 8, height = 5)

# Outlier detection
z_scores <- scale(sales_data$sale.amount)
outliers_z_score <- sales_data[abs(z_scores) > 3, ]

# Additional Visualizations
# 1. Sales Amount Distribution
hist_sales <- ggplot(sales_data, aes(x = sale.amount)) +
  geom_histogram(binwidth = 10, fill = "skyblue", color = "black") +
  ggtitle("Distribution of Sale Amounts") +
  xlab("Sale Amount") +
  ylab("Frequency") +
  theme_minimal()
ggsave("Visuals/histogram_sale_amount.png", hist_sales, width = 8, height = 5)

# Prepare seasonality data
sales_data$date <- as.Date(sales_data$sale.date, format = "%Y-%m-%d")
sales_data$month <- as.numeric(format(sales_data$date, "%m"))
sales_data$season <- ifelse(sales_data$month %in% c(12, 1, 2), "Winter",
                            ifelse(sales_data$month %in% c(3, 4, 5), "Spring",
                                   ifelse(sales_data$month %in% c(6, 7, 8), "Summer", "Fall")))

# Calculate gross margin percentage
sales_data <- sales_data %>%
  mutate(gross_margin_percentage = gross.margin / unit.original.retail)

# 2. Gross Margin by Season (Boxplot)
boxplot_gm_season <- ggplot(sales_data, aes(x = season, y = gross_margin_percentage, fill = season)) +
  geom_boxplot() +
  ggtitle("Gross Margin Percentage by Season") +
  xlab("Season") +
  ylab("Gross Margin %") +
  theme_minimal()
ggsave("Visuals/boxplot_gross_margin_by_season.png", boxplot_gm_season, width = 8, height = 5)

# 3. Seasonal Sales Volume
barplot_sales_volume <- ggplot(sales_data, aes(x = season, fill = season)) +
  geom_bar() +
  ggtitle("Sales Volume by Season") +
  xlab("Season") +
  ylab("Number of Sales") +
  theme_minimal()
ggsave("Visuals/barplot_sales_volume_by_season.png", barplot_sales_volume, width = 8, height = 5)

# 4. Sale Amount vs. Gross Margin (Scatter Plot)
scatter_sales_vs_margin <- ggplot(sales_data, aes(x = sale.amount, y = gross.margin)) +
  geom_point(alpha = 0.5, color = "blue") +
  geom_smooth(method = "lm", color = "red") +
  ggtitle("Sale Amount vs Gross Margin") +
  xlab("Sale Amount") +
  ylab("Gross Margin") +
  theme_minimal()
ggsave("Visuals/scatter_sale_vs_margin.png", scatter_sales_vs_margin, width = 8, height = 5)

# Store-level Visualizations
# 5. Gross Margin by Store
gross_margin_by_store <- sales_data %>%
  group_by(store, city) %>%
  summarise(
    total_sale_amount = sum(sale.amount, na.rm = TRUE),
    total_ext_cost = sum(ext.cost, na.rm = TRUE),
    blended_gross_margin = (total_sale_amount - total_ext_cost) / total_sale_amount * 100
  )
write.csv(gross_margin_by_store, "Outputs/gross_margin_by_store.csv", row.names = FALSE)

barplot_gm_store <- ggplot(gross_margin_by_store, aes(x = city, y = blended_gross_margin, fill = city)) +
  geom_bar(stat = "identity") +
  ggtitle("Blended Gross Margin by Store") +
  xlab("Store (City)") +
  ylab("Blended Gross Margin (%)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
ggsave("Visuals/barplot_gross_margin_by_store.png", barplot_gm_store, width = 8, height = 5)

# 6. Sales Volume by Store
barplot_sales_volume_store <- ggplot(sales_data, aes(x = city, fill = city)) +
  geom_bar() +
  ggtitle("Sales Volume by Store") +
  xlab("Store (City)") +
  ylab("Number of Sales") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
ggsave("Visuals/barplot_sales_volume_by_store.png", barplot_sales_volume_store, width = 8, height = 5)

# 7. Gross Margin by Season and Store (Boxplot)
boxplot_store_season <- ggplot(sales_data, aes(x = season, y = gross_margin_percentage, fill = city)) +
  geom_boxplot() +
  ggtitle("Gross Margin % by Season and Store") +
  xlab("Season") +
  ylab("Gross Margin %") +
  theme_minimal()
ggsave("Visuals/boxplot_gm_by_store_and_season.png", boxplot_store_season, width = 8, height = 5)

# Hypothesis Testing (ANOVA)
anova_result <- aov(gross_margin_percentage ~ season, data = sales_data)
print(summary(anova_result))
anova_p_value <- summary(anova_result)[[1]]$`Pr(>F)`[1]

if (anova_p_value < 0.05) {
  print("Reject the null hypothesis: There is a significant difference in GM% across the seasons.")
} else {
  print("Fail to reject the null hypothesis: There is no significant difference in GM% across the seasons.")
}

# Regression Analysis
price_category_dummies <- model.matrix(~ price.category - 1, data = sales_data)
category_dummies <- model.matrix(~ category - 1, data = sales_data)
interaction_terms_price <- sales_data$sale.amount * price_category_dummies
interaction_terms_category <- sales_data$sale.amount * category_dummies
sales_data <- cbind(sales_data, interaction_terms_price, interaction_terms_category)

model <- lm(gross.margin ~ sale.amount + ext.cost + category + price.category +
              loyalty.member + interaction_terms_price + interaction_terms_category,
            data = sales_data)
summary_model <- summary(model)
coefficients_table <- summary_model$coefficients
write.csv(coefficients_table, "Outputs/model_summary.csv")
write_xlsx(as.data.frame(coefficients_table), "Outputs/model_summary.xlsx")
