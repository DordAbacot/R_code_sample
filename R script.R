library(ggplot2)
library(dplyr)
library(tidyr)
install.packages("openxlsx")
library(tidyverse)
library(openxlsx)

op<-read.csv("CRTW operational and beta items.csv")
op$Learning.Outcome <- substr(op$Learning.Outcome, 1, 2)


# A basic scatterplot with color depending on Species
ggplot(op, aes(x=op$Item.Difficulty, y=op$Item.Discrimination, color=Item.Classification)) + 
  geom_point(size=2) + 
  labs(x = "Item Difficulty", y = "Item Discrimination")+
  geom_vline(xintercept =  0.95, linetype="dotted", color = "black", size=1) + 
  geom_vline(xintercept =  0.25, linetype="dotted", color = "black", size=1) + 
  geom_hline(yintercept = 0) + 
  xlim(0.15, 1) + ylim(-0.3, 0.9) + 
  scale_colour_brewer(palette = "Set1") + 
  geom_rect(aes(xmin = 0, xmax = 0, ymin = 0, ymax = 0),
            alpha = 1/5) +
  scale_color_manual(values = c("black", "red"))+
  theme(panel.background = element_rect(fill = "#F5F5F5"),
        panel.grid.major = element_line(color = "#FFFFFF"),
        panel.grid.minor = element_line(color = "#FFFFFF"))


#creating histogram
dens <- density(op$Item.Difficulty[op$Item.Classification == "Operational"])
hist(op$Item.Difficulty[op$Item.Classification == "Operational"], main = "Item Difficulty Distribution", col= "#F0FFFF", xlab = "", freq = F)
lines(dens, lwd = 2, col = "red")

plot(dens, type = "l", lwd = 2, col = "red", xlab = "Item Difficulty", ylab = "Density", main = "Operational Item Difficulty Distribution")

#----Boxplots per learning outcome----

boxplot(Item.Difficulty~Learning.Outcome,
        data = op,
        col = c("#e2ac99", "#dca690","#d6a187","#d09b7e","#cb9575","#c58f6b","#bf8a62","#b98459", "#b37e50"),
        las = 1,
        xlab = "Learning Outcome",
        ylab = " Item Difficulty",
        cex.names = 0.1)





#----creating the Distractor Analysis matrices----
da <- read.csv("all_items_da.csv")
da <- da[, -c(2, 8,10,12, 4, 11, 6, 3)]

abilityDivision <- 3 #number to divide the candidates by ability groups

da$candidateAnswer <- gsub("[.| ]", "", da$candidateAnswer) #cleaning string
da$candidateAnswer <- ifelse(da$candidateAnswer == "", "Omitted", da$candidateAnswer) #if blank then omitted
da$keyAnswer <- gsub("[.| ]", "", da$keyAnswer) #cleaning string
da$Quartiles <- ntile(da$Countcorrect, abilityDivision)
da$Quartiles <- paste("Ability", da$Quartiles, sep = " ")
da$QID <- gsub("[ ]", "", da$QID)

# Define all possible options
all_options <- c("A", "B", "C", "D", "Omitted")

# Create a list of data frames (each being a DA matrix) 
list_of_dfs <- da %>%
  group_by(QID, candidateAnswer, Quartiles) %>%
  summarise(count = n(), .groups = "drop") %>%
  ungroup() %>%
  complete(QID, candidateAnswer = all_options, Quartiles, fill = list(count = 0)) %>%
  pivot_wider(names_from = Quartiles, values_from = count, values_fill = 0) %>%
  group_by(QID) %>%
  group_split()

#mark Key Answer
for (i in seq_along(list_of_dfs)) {
  for (l in 1:5) {
    ifelse(list_of_dfs[[i]][l,2]== da$keyAnswer[which(da$QID == as.character(list_of_dfs[[i]][1,1]))], 
           list_of_dfs[[i]][l,2] <- paste0(list_of_dfs[[i]][l,2], "*"), 
           list_of_dfs[[i]][l,2] <- list_of_dfs[[i]][l,2])
  }
}

# Initialize a new workbook
wb <- createWorkbook()

# Add a worksheet
addWorksheet(wb, sheetName = "DAMatrices")

# Variables to store the current row and column for writing data
current_row <- 1
current_col <- 1

# Format for header
style_header <- createStyle(textDecoration = "bold", border = "TopBottomLeftRight")  # Define a style with borders

# Add each DA matrix to the workbook
for(i in seq_along(list_of_dfs)) {
  # Remove QID column and keep candidateAnswer for options
  df_to_write <- list_of_dfs[[i]]
  df_to_write$QID <- NULL
  
  # Replace "candidateAnswer" with the item name
  colnames(df_to_write)[1] <- pull(list_of_dfs[[i]], QID)[1]
  
  # Separate headers and data
  headers <- matrix(colnames(df_to_write), byrow = FALSE, nrow = 1, ncol = length(df_to_write), dimnames = NULL)
  df_to_write <- as.data.frame(lapply(df_to_write, as.character), stringsAsFactors = FALSE)  # Convert to character to align all data
  df_to_write_no_headers <- matrix(as.matrix(df_to_write), ncol = ncol(df_to_write), dimnames = NULL)  # Data without headers
  colnames(df_to_write_no_headers) <- headers #replace headers to correct for spacing
  
  # Add style to header and first column
  addStyle(wb, sheet = "DAMatrices", style = style_header, rows = current_row, cols = current_col:length(headers))
 
  # Write data with borders
  writeData(wb, sheet = "DAMatrices", x = df_to_write_no_headers, startRow = current_row, startCol = current_col, borders = "all")
  
  # Update the current row (add 2 for some spacing between matrices)
  current_row <- current_row + nrow(df_to_write) + 2

  QID_name<-headers[1 ,1]
  
  #Plotting
  df_to_plot <- as.data.frame(df_to_write_no_headers)
  df_to_plot[, 1]<-df_to_plot[, 1]
  levels <- unique(df_to_plot[, 1])
  
  # Create a named vector for the colors
  colors_group <- c("Blue", "Green", "Purple", "Brown", "Gray")
  names(colors_group) <- levels
  #[!grepl("\\*", levels)]
  
  # Assign colors to the 'Color' column based on QID_name
  df_to_plot$Color <- colors_group[df_to_plot[, 1]]
  df_to_plot$Color <- ifelse(grepl("\\*", df_to_plot[, 1]), "Red", df_to_plot$Color)
  #df_to_plot$Color[grepl("\\*", df_to_plot[, 1])] <- "Red"
  
  # Converting to long-form data frame
  df_to_plot <- reshape2::melt(df_to_plot, id.vars = c(QID_name, "Color"))
  df_to_plot$value <- as.numeric(df_to_plot$value) # values from character to numeric
  color_mapping <- setNames(df_to_plot$Color, df_to_plot[, 1]) #colors with their names
  
  #plotting
  matrix_plot<-ggplot(df_to_plot, aes(x = variable, y = value, group = df_to_plot[, 1])) +
    geom_line(aes(color = df_to_plot[, 1]), size = 1.25) +
    geom_point(color = color_mapping, size = 3) +
    scale_color_manual(values = color_mapping) +
    labs(x = "", y = "", color = "Option", title = QID_name) +
    theme_bw() +
    theme(legend.position = "right", 
          legend.background = element_rect(fill="gray90", linewidth = 5, linetype="dotted"))
  
  #saving as JPG
  new_folder <- "DA graphs/"
  dir.create(new_folder, showWarnings = FALSE)
  ggsave(paste0(new_folder,QID_name,".jpg"), plot = matrix_plot, width = 6, height = 4, dpi = 300)
}


saveWorkbook(wb, "DAMatrices.xlsx", overwrite = TRUE)

