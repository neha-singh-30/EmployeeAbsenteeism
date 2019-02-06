#First clean the environment
rm(list = ls())

#set working directory
setwd("/home/neha/home/Project_2")

#Load required packages
x = c("xlsx", "DMwR","corrgram", "caret", "usdm", "rpart", "DataCombine", "randomForest","e1071", "ggplot2", "inTrees", "lsr")

lapply(x, require, character.only=TRUE)

rm(x)

#Load the data
data = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1)

########Explore the data#########

str(data)
dim(data)
summary(data)
class(data)
colnames(data)

#convert the variables into their respective types
data$ID = as.factor(as.character(data$ID))
data$Reason.for.absence = as.factor(as.character(data$Reason.for.absence))
data$Month.of.absence = as.factor(as.character(data$Month.of.absence))
data$Day.of.the.week = as.factor(as.character(data$Day.of.the.week))
data$Seasons = as.factor(as.character(data$Seasons))
data$Disciplinary.failure = as.factor(as.character(data$Disciplinary.failure))
data$Education = as.factor(as.character(data$Education))
data$Social.drinker = as.factor(as.character(data$Social.drinker))
data$Social.smoker = as.factor(as.character(data$Social.smoker))

##########Missing Value Analysis##########

missing_val = data.frame(apply(data,2,function(x)sum(is.na(x))))
missing_val$Columns = row.names(missing_val)
row.names(missing_val) = NULL
names(missing_val)[1] = "Missing_percentage"
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(data))*100
missing_val = missing_val[order(-missing_val$Missing_percentage),]
missing_val = missing_val[, c(2,1)]

#store the missing value information
write.csv(missing_val, "Missing_Value_Percentage", row.names = F)

#Now let's impute missing value
#First we will create missing value in one cell to check

data$Body.mass.index[10] #value of this cell is 29
data$Body.mass.index[10] = NA

#Actual value : 29
#Mean value : 26.68
#Median Value : 25
#KNN Value : 29

#Now will apply mean method to impute this generated value and observe the result

#Mean Method
#data$Body.mass.index[is.na(data$Body.mass.index)] = mean(data$Body.mass.index, na.rm = T)

#Median Method
#data$Body.mass.index[is.na(data$Body.mass.index)] = median(data$Body.mass.index, na.rm = T)

#KNN Method
data = knnImputation(data, k =3)

#after applying mean, median and KNN, we found KNN method is more accurate, hence we freeze this
#check for missing value
sum(is.na(data)) #there is no missing value now

#########Outlier Analysis##############

numeric_index = sapply(data, is.numeric)
numeric_data = data[, numeric_index]

#Store all the column names excluding target variable name
cnames = colnames(numeric_data)[-12]

#Plotting boxplot to detect outliers

for (i in 1:length(cnames)) {
  assign(paste0("gn",i), ggplot(aes_string( y = (cnames[i]), x= "Absenteeism.time.in.hours") , data = subset(data)) +
           stat_boxplot(geom = "errorbar" , width = 0.5) +
           geom_boxplot(outlier.color = "red", fill = "blue", outlier.shape = 20, outlier.size = 1, notch = FALSE)+
           theme(legend.position = "bottom")+
           labs(y = cnames[i], x= "Absenteeism.time.in.hours")+
           ggtitle(paste("Boxplot" , cnames[i])))
  #print(i)
}

options(warn = 0)

#lets plot the boxplots
gridExtra::grid.arrange(gn1, gn2,gn3, ncol=3)
gridExtra::grid.arrange(gn4,gn5,gn6, ncol=3)
gridExtra::grid.arrange(gn7,gn8,gn9, ncol =3)
gridExtra::grid.arrange(gn10,gn11, ncol =3 )

#getting outliers using boxplot.stat method
for (i in cnames) {
  print(i)
  val = data[,i][data[,1] %in% boxplot.stats(data[,i])$out]
  print(length(val))
  print(val)
}

#Make each outlier as NA
for (i in cnames) {
  val = data[,i][data[,i] %in% boxplot.stats(data[,i])$out]
  data[,i][data[,i] %in% val] = NA
}

#checking the missing values
sum(is.na(data)) 

#Impute the values using KNN imputation method
data = knnImputation(data, k=3)

#Check again for missing value if present in case
sum(is.na(data))

#############Feature Selection################

#Correlation plot

corrgram(data[, cnames],order = F, upper.panel = panel.pie, text.panel = panel.txt, main = "correlation plot" )

#ANOVA test

anova_test = aov(Absenteeism.time.in.hours ~ ID + Day.of.the.week + Education + Social.smoker + Social.drinker + Reason.for.absence + Seasons + Month.of.absence + Disciplinary.failure, data = data)
summary(anova_test)

#Dimensionality Reduction

data = subset(data,select = -c(Weight, Education, Social.drinker, Seasons, Month.of.absence, Disciplinary.failure))

############Feature Scaling##############

# Using historam to check how if our data is normally disrtibuted or not
hist(data$Body.mass.index)
hist(data$Age)
hist(data$Absenteeism.time.in.hours)
hist(data$Transportation.expense)
hist(data$Work.load.Average.day.)

#Hence we will choose normalisation instead of standardisation bcz variables are not normally distributed

num_names = colnames(data[,sapply(data,is.numeric)])
num_names = num_names[-11]

for (i in num_names) {
  print(i)
  data[,i] = (data[,i] - min(data[,i]))/(max(data[,i]) - min(data[,i]))
}

###########Model Development#############

rmExcept("data")

#1.Decision Tree Regression

#devide the data into train and test
train_index = sample(1:nrow(data), 0.8*nrow(data))
train = data[train_index,]
test = data[-train_index,]

#Model
DT_Reg = rpart(Absenteeism.time.in.hours ~. , data = train, method = "anova")

#Lets predict for test cases
Predictions_DT = predict(DT_Reg, test[-17])

#Evaluate the performance of model, using rmse as the data is time series data
Rmse_DT = regr.eval(test[,15], Predictions_DT, stats = 'rmse')

#RMSE Value : 10.98
#Accuracy : 89.02

#2.Random Forest Regression

RF_Reg = randomForest(Absenteeism.time.in.hours ~. , train, importance = TRUE, ntree = 100)

#Extract rules fromn random forest
#transform rf object to an inTrees' format
treeList = RF2List(RF_Reg)

#Extract rules
exec = extractRules(treeList, train[-15])

#Visualize some rules
exec[1:2,]

#Make rules more readable
ReadableRules = presentRules(exec, colnames(train))
ReadableRules[1:2,]

#Get rule metrics
RuleMetric = getRuleMetric(exec, train[-15], train$Absenteeism.time.in.hours)
RuleMetric[1:2,]

#Predict test data using random forest model
Predictions_RF = predict(RF_Reg, test[-15])

##Evaluate the performance of model
Rmse_RF = regr.eval(test[,15], Predictions_RF, stats = 'rmse')

#RMSE Value : 8.53
#Hence Accuracy : 91.47

#3.Linear Regression

#First we need to check for multicollinearity

#Removing categorical data for checking multicollinearity
LR_data = subset(data, select = -c(ID, Reason.for.absence, Month.of.absence, Day.of.the.week, Social.smoker, Social.drinker))
vif(LR_data[,-11])
vifcor(LR_data[,-11], th=0.9)

#Now will run the model
LR_Model = lm(Absenteeism.time.in.hours ~. , data = train[, !colnames(train) %in% c("ID")])

#Summary of the model
summary(LR_Model)

#Predict
Predictions_LR = predict(LR_Model, test[,1:14])

#Calculate RMSE
RMSE(test[,15], Predictions_LR)

#RMSE Value : 9.42
#Accuracy : 90.58




############Visvualization############

#library(scales)
#library(psych)
# 
# #Reasons for absence count, 23 is highest of all
# ggplot(data, aes_string(x=data$Reason.for.absence)) + geom_bar(stat = "count", fill ="blue") +
#  xlab("Reason") + ylab("Count") + ggtitle("Distribution")+ theme(text = element_text(size = 15))
# 
# #Pet count
# ggplot(data, aes_string(x=data$Pet)) + geom_bar(stat = "count", fill = "Darkslateblue") +
#   xlab("No. of Pets") + ylab("Count") + ggtitle("Pet Distribution")
# 
# plot(Absenteeism.time.in.hours ~ Pet , data = data)
# #people with atleast one pet show less absentism
# 
# #Transportation expense
#ggplot(data , aes_string(x=data$Transportation.expense)) + geom_bar(stat = "count", fill = "DarkslateBlue") + xlab("Transportation expense") +
#   ylab("Count") + scale_y_continuous(breaks=pretty_breaks(n=10)) + ggtitle("Transportation expanse distribution") + theme(text=element_text(size=15))
# 
# plot(Absenteeism.time.in.hours ~ Transportation.expense, data = data)