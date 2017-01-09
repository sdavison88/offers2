library(arules)
library(arulesViz)
library(stringr)
library(splitstackshape)
library(plyr)
library(lubridate)
library(xlsx)

# import & format data
nna_basket_data <- read.csv("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/data/nna_union_data_12_2_16_v2.csv")
nna_basket_data$dmi_opcodes <- as.character(nna_basket_data$dmi_opcodes)
nna_basket_data$Coupons <- as.character(nna_basket_data$Coupons)
nna_basket_data$ro_close_date <- as.Date(nna_basket_data$ro_close_date)

# relabel segments
nna_basket_data$segment <- factor(nna_basket_data$segment,
                                  levels = c(1, 2, 3, 4, 5),
                                  labels = c("New", "Retained", "At-Risk", "Recapture", "8+ Years"))

# collapse onto vin level for coupon basket analysis
nna_basket_data_collapsed <- ddply(nna_basket_data, .(vehicle_id), summarize, Coupons = toString(Coupons))
nna_basket_data_collapsed$segment <- nna_basket_data[match(nna_basket_data_collapsed$vehicle_id, nna_basket_data$vehicle_id), "segment"]

# import op code & coupon dictionaries
op_code_dict <- read.csv("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/data/service_dictionary.csv")
coupon_dict <- read.csv("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/data/coupon_dict.csv")
coupon_type <- read.csv("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/data/coupon_cost_lookup.csv")

# create target for modeling
nna_basket_data$total_pay <- as.numeric(nna_basket_data$labor_amt + nna_basket_data$parts_amt + nna_basket_data$misc_amt)
nna_basket_data$total_pay[which(nna_basket_data$total_pay < 0)] <- 0

# create campaign_period split for modeling
nna_basket_data$month <- month(nna_basket_data$ro_close_date)
nna_basket_data$campaign_period <- NA
nna_basket_data$campaign_period[which(nna_basket_data$month %in% c(2, 3))] <- 1
nna_basket_data$campaign_period[which(nna_basket_data$month %in% c(5, 6))] <- 2
nna_basket_data$campaign_period[which(nna_basket_data$month %in% c(8, 9))] <- 3
nna_basket_data$campaign_period[which(nna_basket_data$month %in% c(11, 12))] <- 4

# confirm all campaigns are represented
table(nna_basket_data$campaign_period)

# remove observations without a campaign period
nna_basket_data <- nna_basket_data[is.na(nna_basket_data$campaign_period) == FALSE,]

# define splits for modeling & basket analysis
segs <- sort(unique(nna_basket_data$segment))
campaign_period <- sort(unique(nna_basket_data$campaign_period))

###################################
##### BASKET ANALYSIS COUPONS #####
###################################

# NOTE: RUN AT THE VIN LEVEL

basket_output_coupons <- data.frame()
for (i in segs) {
  temp_table <- subset(nna_basket_data, segment == i, select = c(vehicle_id, Coupons))
  write(temp_table$Coupons, "nna_arules")
  temp_tcs <- read.transactions("nna_arules", format = "basket", sep = ",", rm.duplicates = TRUE)
  #temp_tcs@itemInfo$labels <- coupon_dict[match(temp_tcs@itemInfo$labels, coupon_dict$Coupon_Code), 2]
  basket_rules <- apriori(temp_tcs, parameter = list(target = "rules", sup = 0.00005))
  # save images for each segment
  # png(filename = paste("/Users/elliot.henry/Desktop/NNA/basket_coupon_plot_", i, "_", Sys.Date(), ".png", sep = ""), bg = "transparent")
  # plot(basket_rules, method = "graph", main = paste("Coupon Association Rules for ", i, " Segment", sep = ""), control = list(type = "items"))
  # dev.off()
  # add to initialized matrix
  temp_rules <- as(basket_rules, "data.frame")
  temp_rules$Segment <- i
  basket_output_coupons <- rbind(basket_output_coupons, temp_rules)
}

# convert to character
basket_output_coupons$rules_description <- as.character(basket_output_coupons$rules)
coupon_dict$Coupon_Code <- as.character(coupon_dict$Coupon_Code)
coupon_dict$Coupon_Name <- as.character(coupon_dict$Coupon_Name)

# replace coupon codes with coupon descriptions
for (i in 1:nrow(coupon_dict)) {
  basket_output_coupons$rules_description <- gsub(as.character(coupon_dict[i,1]), as.character(coupon_dict[i,2]), basket_output_coupons$rules_description)
}

# remove any express service rules
basket_output_coupons <- basket_output_coupons[grep("Express", basket_output_coupons$rules_description, ignore.case = FALSE, invert = TRUE),]
basket_output_coupons <- basket_output_coupons[grep("ES", basket_output_coupons$rules_description, ignore.case = FALSE, invert = TRUE),]
basket_output_coupons <- basket_output_coupons[grep("Exp Srv", basket_output_coupons$rules_description, ignore.case = FALSE, invert = TRUE),]

# separate out LH and RH rules
basket_output_coupons$LHS <- gsub("=>.*$", "", basket_output_coupons$rules_description)
basket_output_coupons$RHS <- gsub(".*=> ", "", basket_output_coupons$rules_description)

# only include rules with lift > 1
basket_output_coupons <- basket_output_coupons[basket_output_coupons$lift > 1,]

# reorder columns
basket_output_coupons <- basket_output_coupons[c("Segment", "rules", "rules_description", "LHS", "RHS", "support", "confidence", "lift")]

# save file
#write.table(basket_output_coupons, file = paste("/Users/elliot.henry/Desktop/NNA/basket_coupons_", Sys.Date(), ".csv", sep = ""), sep = ",", row.names = FALSE)

####################################
##### BASKET ANALYSIS OP CODES #####
####################################

# NOTE: RUN AT THE SERVICE LEVEL

basket_output_op_codes <- data.frame()
for (i in segs) {
  temp_table <- subset(nna_basket_data, segment == i, select = c(service_id, dmi_opcodes))
  write(temp_table$dmi_opcodes, "nna_arules")
  temp_tcs <- read.transactions("nna_arules", format = "basket", sep = ",", rm.duplicates = TRUE)
  temp_tcs@itemInfo$labels <- op_code_dict[match(temp_tcs@itemInfo$labels, op_code_dict$dmi_opcode), 2]
  basket_rules <- apriori(temp_tcs, parameter = list(target = "rules", sup = 0.01))
  # save images for each segment
  # png(filename = paste("/Users/elliot.henry/Desktop/NNA/basket_op_code_plot_", i, "_", Sys.Date(), ".png", sep = ""), bg = "transparent")
  # plot(basket_rules, method = "graph", main = paste("Op Code Association Rules for ", i, " Segment", sep = ""), control = list(type = "items"))
  # dev.off()
  # add to initialized matrix
  temp_rules <- as(basket_rules, "data.frame")
  temp_rules$Segment <- i
  basket_output_op_codes <- rbind(basket_output_op_codes, temp_rules)
}

# separate out LH and RH rules
basket_output_op_codes$LHS <- gsub("=>.*$", "", basket_output_op_codes$rules)
basket_output_op_codes$RHS <- gsub(".*=> ", "", basket_output_op_codes$rules)

# only include rules with lift > 1
basket_output_op_codes <- basket_output_op_codes[basket_output_op_codes$lift > 1,]

basket_output_op_codes <- basket_output_op_codes[c("Segment", "rules", "LHS", "RHS", "support", "confidence", "lift")]

# save file
#write.table(basket_output_op_codes, file = paste("/Users/elliot.henry/Desktop/NNA/basket_op_codes_", Sys.Date(), ".csv", sep = ""), sep = ",", row.names = FALSE)

###########################
##### COUPON MODELING #####
###########################

# create dummy variables
model_df <- cSplit_e(nna_basket_data, "Coupons", ",", type = "character", fill = 0)

# summarize the data and append onto output table
coupon_summary <- data.frame()
rnum <- 0
for (j in campaign_period) {
  for (n in segs) {
    for (i in names(model_df)[(ncol(nna_basket_data)+1):(ncol(model_df))]) {
      rnum <- rnum + 1
      coupon_summary[rnum, "campaign_period"] <- j
      coupon_summary[rnum, "segment"] <- n
      coupon_summary[rnum, "coupon"] <- i
      coupon_summary[rnum, "appearance_count"] <- sum(model_df[model_df$campaign_period == j & model_df$segment == n,i])
      coupon_summary[rnum, "appearance_percent"] <- mean(model_df[model_df$campaign_period == j & model_df$segment == n,i])
    }
  }
}
coupon_summary$coupon_name <- gsub(".*Coupons_", "", coupon_summary$coupon)

# modeling
model_summary <- data.frame()
for (i in campaign_period) {
  for (n in segs) {
    temp_df <- model_df[model_df$campaign_period == i & model_df$segment == n,]
    drops <- vector()
    for (j in names(temp_df)[(ncol(nna_basket_data)+1):(ncol(temp_df))]) {
      if (sum(temp_df[,j]) < 20) { # drops any columns with < 20 observations
        drops <- c(drops, j)
        }
      }
    temp_df <- temp_df[, !colnames(temp_df) %in% drops]
    temp_df[(ncol(nna_basket_data)+1):(ncol(temp_df))] <- lapply(temp_df[(ncol(nna_basket_data)+1):(ncol(temp_df))], as.factor)
    temp_glm <- glm(as.formula(paste("total_pay~", paste(names(temp_df)[(ncol(nna_basket_data)+1):(ncol(temp_df))], collapse = "+"), "-1")), data = temp_df, family = "gaussian")
    temp_coupon_names <- rownames(summary(temp_glm)$coefficients)
    temp_glm_summary <- as.data.frame(summary(temp_glm)$coefficients)
    rownames(temp_glm_summary) <- NULL
    temp_glm_summary$coupon_name <- temp_coupon_names
    temp_glm_summary$campaign_period <- i
    temp_glm_summary$segment <- n
    model_summary <- rbind(model_summary, temp_glm_summary)
    #rm(temp_df, temp_glm, temp_coupon_names, temp_glm_summary)
  }
}

# remove model intercepts
model_summary <- model_summary[substr(model_summary$coupon_name, nchar(model_summary$coupon_name), nchar(model_summary$coupon_name)) != 0,]
model_summary$coupon_name <- substr(model_summary$coupon, 1, nchar(model_summary$coupon)-1)
model_summary$coupon_name <- gsub(".*Coupons_", "", model_summary$coupon_name)

# merge model summary and coupon summary tables
model_output <- merge(model_summary, coupon_summary, by.x = c("coupon_name", "campaign_period", "segment"), by.y = c("coupon_name","campaign_period", "segment"))
colnames(model_output)[7] <- "pvalue"
model_output <- model_output[,-c(8)]

# only look at significant p values and where the coefficient is positive
model_output <- model_output[model_output$pvalue < 0.05,]

# add coupon descriptions
model_output$coupon_desc <- coupon_dict[match(model_output$coupon_name, coupon_dict$Coupon_Code),2]
model_output$service_category <- coupon_dict[match(model_output$coupon_name, coupon_dict$Coupon_Code),3]

# add coupon type and amount off
model_output$coupon_type <- coupon_type[match(model_output$coupon_name, coupon_type$Coupon_Code),3]
model_output$coupon_price <- coupon_type[match(model_output$coupon_name, coupon_type$Coupon_Code),4]

# add value added and ranking
model_output$Value_Added <- sin(model_output$appearance_count)^2 * model_output$Estimate
model_output$Rank <- rank(-model_output$Value_Added, ties.method = "first")

# convert promo periods to text
model_output$campaign_period <- factor(model_output$campaign_period,
                                  levels = c(1, 2, 3, 4),
                                  labels = c("Feb-Mar", "May-Jun", "Aug-Sep", "Nov-Dec"))

# remove express service
model_output <- model_output[grep("Express", model_output$coupon_desc, ignore.case = FALSE, invert = TRUE),]
model_output <- model_output[grep("ES", model_output$coupon_desc, ignore.case = FALSE, invert = TRUE),]
model_output <- model_output[grep("Exp Srv", model_output$coupon_desc, ignore.case = FALSE, invert = TRUE),]

# reorder columns
model_output <- model_output[c("campaign_period", "segment", "coupon_name", "coupon_desc", "service_category", "Rank", "Value_Added", "coupon_type", "coupon_price", "appearance_count", "appearance_percent", "Estimate", "Std. Error", "t value", "pvalue")]

# save file
#write.table(basket_output_op_codes, file = paste("/Users/elliot.henry/Desktop/NNA/coupon_model_output_", Sys.Date(), ".csv", sep = ""), sep = ",", row.names = FALSE)

# save all output to one excel file
write.xlsx(basket_output_coupons, file = paste("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/output/Analytics_Rankings_",Sys.Date(),".xlsx", sep = ""), sheetName = "Coupon_Associations", row.names = FALSE)
write.xlsx(basket_output_op_codes, file = paste("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/output/Analytics_Rankings_",Sys.Date(),".xlsx", sep = ""), sheetName = "Op_Code_Associations", append = TRUE, row.names = FALSE)
write.xlsx(model_output, file = paste("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/output/Analytics_Rankings_",Sys.Date(),".xlsx", sep = ""), sheetName = "Coupon_Rankings", append = TRUE, row.names = FALSE)
write.xlsx(coupon_type, file = paste("/Users/elliot.henry/Desktop/NNA/fy17_coupon_analysis/output/Analytics_Rankings_",Sys.Date(),".xlsx", sep = ""), sheetName = "Coupon_Pricing_Lookup", append = TRUE, row.names = FALSE)



