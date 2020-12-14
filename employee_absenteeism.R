# Remove all the objects stored
rm(list = ls())

#Setting working directory
setwd("E:/edWisor/Project")

#Loading libraries
x = c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced","dummies", "e1071", "Information",
      "MASS", "rpart","gbm", "ROSE","sampling", "DataCombine", "xlsx","inTrees","usdm")
#Install packages
lapply(x,require, character.only = TRUE)

rm(x)

## Read the data
emp_abs = read.xlsx("Absenteeism_at_work_Project.xls",sheetIndex=1, header = T)

#####################################Explore the data####################################
dim(emp_abs)
names(emp_abs)
head(emp_abs,5)
str(emp_abs)
summary(emp_abs)

#Numeric to factor transformation
emp_abs$ID = as.factor(as.character(emp_abs$ID))
emp_abs$Reason.for.absence[emp_abs$Reason.for.absence %in% 0]=20
emp_abs$Reason.for.absence = as.factor(as.character(emp_abs$Reason.for.absence))
emp_abs$Month.of.absence[emp_abs$Month.of.absence %in% 0]= NA
emp_abs$Month.of.absence = as.factor(as.character(emp_abs$Month.of.absence))
emp_abs$Day.of.the.week= as.factor(as.character(emp_abs$Day.of.the.week))
emp_abs$Seasons = as.factor(as.character(emp_abs$Seasons))
emp_abs$Disciplinary.failure = as.factor(as.character(emp_abs$Disciplinary.failure))
emp_abs$Education = as.factor(as.character(emp_abs$Education))
emp_abs$Son = as.factor(as.character(emp_abs$Son))
emp_abs$Social.drinker = as.factor(as.character(emp_abs$Social.drinker))
emp_abs$Social.smoker = as.factor(as.character(emp_abs$Social.smoker))
emp_abs$Pet = as.factor(as.character(emp_abs$Pet))

str(emp_abs)

############################# Missing value analysis ###########################
missing_val = data.frame(apply(emp_abs,2,function(x){sum(is.na(x))}))

#Converting row names into column names
missing_val$Columns = row.names(missing_val)
row.names(missing_val) = NULL
# Rename the variable
names(missing_val)[1]= "Missing_percentage"
# Calculating percentage
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(emp_abs))*100
# Arrange in descending order
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL
# Rearrange the column
missing_val= missing_val[,c(2,1)]

# Writing output result back to disk
write.csv(missing_val,"emp_abs_mp.csv",row.names = F)

# Visualization of missing value
ggplot(data = missing_val[1:6,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
   geom_bar(stat = "identity",fill = "greenyellow")+xlab("Parameter")+
   ggtitle("Missing data percentage (Train)") + theme_bw()

# Identifying imputation method (mean, median, KNN)
emp_abs[["Body.mass.index"]][5]= NA

# Actual value = 30
# Mean imputed value = 26.67938
# Median imputed value = 25
# KNN imputed value = 30

# Mean method
#emp_abs$Body.mass.index[is.na(emp_abs$Body.mass.index)] = mean(emp_abs$Body.mass.index, na.rm = T)

# Median method
#emp_abs$Body.mass.index[is.na(emp_abs$Body.mass.index)] = median(emp_abs$Body.mass.index, na.rm = T)

# KNN imputation
emp_abs = knnImputation(emp_abs, k=3)
sum(is.na(emp_abs))

write.csv(emp_abs, 'emp_abs_missing.csv', row.names = F)
##################################### Graphs ###################################

b1<-ggplot(emp_abs, aes_string(x=reorder(emp_abs$ID),fill =emp_abs$ID)) + geom_bar(stat ='count')+ ggtitle("Count of ID") + theme_bw()
b2<-ggplot(emp_abs, aes_string(reorder(emp_abs$Reason.for.absence),fill=emp_abs$Reason.for.absence)) + geom_bar(stat = "count")+ ggtitle("Count of Reason.for.absence") + theme_bw()
b3<-ggplot(emp_abs, aes_string(reorder(emp_abs$Month.of.absence),fill=emp_abs$Month.of.absence)) + geom_bar(stat = "count")+ ggtitle("Count of Month.of.absence") + theme_bw()
b4<-ggplot(emp_abs, aes_string(reorder(emp_abs$Day.of.the.week),fill=emp_abs$Day.of.the.week)) + geom_bar(stat = "count")+ ggtitle("Count of Day of the week") + theme_bw()
b5<-ggplot(emp_abs, aes_string(reorder(emp_abs$Seasons),fill=emp_abs$Seasons)) + geom_bar(stat = "count")+ ggtitle("Count of Seasons") + theme_bw()
b6<-ggplot(emp_abs, aes_string(reorder(emp_abs$Disciplinary.failure),fill=emp_abs$Disciplinary.failure)) + geom_bar(stat = "count")+ ggtitle("Count of Disciplinary.failure") + theme_bw()
b7<-ggplot(emp_abs, aes_string(reorder(emp_abs$Education),fill=emp_abs$Education)) + geom_bar(stat = "count")+ ggtitle("Count of Education") + theme_bw()
b8<-ggplot(emp_abs, aes_string(reorder(emp_abs$Son),fill=emp_abs$Son)) + geom_bar(stat = "count")+ ggtitle("Count of Son") + theme_bw()
b9<-ggplot(emp_abs, aes_string(reorder(emp_abs$Social.drinker),fill=emp_abs$Social.drinker)) + geom_bar(stat = "count")+ ggtitle("Count of Social drinker") + theme_bw()
b10<-ggplot(emp_abs, aes_string(reorder(emp_abs$Social.smoker),fill=emp_abs$Social.smoker)) + geom_bar(stat = "count")+ ggtitle("Count of Social.smoker") + theme_bw()
b11<-ggplot(emp_abs, aes_string(reorder(emp_abs$Pet),fill=emp_abs$Pet)) + geom_bar(stat = "count")+ ggtitle("Count of Pet") + theme_bw()

b1
b2
gridExtra::grid.arrange(b3,b4,ncol=2)
gridExtra::grid.arrange(b5,b6,ncol=2)
gridExtra::grid.arrange(b7,b8,ncol=2)
gridExtra::grid.arrange(b9,b10,b11,ncol=3)


################################ Outlier analysis ##################################
# ## Box Plots - Distribution and Outlier Check
numeric_index = sapply(emp_abs,is.numeric) #selecting only numeric
numeric_data = emp_abs[,numeric_index]
cnames = colnames(numeric_data)

# Selecting categorical data
categorical_data = emp_abs[,!numeric_index]

# 
for(i in 1:ncol(numeric_data)) {
  assign(paste0("gn",i),ggplot(data = emp_abs, aes_string(y = numeric_data[,i])) +
  stat_boxplot(geom = "errorbar", width = 0.5) +
  geom_boxplot(outlier.colour = "red", fill = "grey", outlier.size = 1) +
  labs(y = colnames(numeric_data[i])) +
  ggtitle(paste("Boxplot: ",colnames(numeric_data[i]))))
}
# 
# Arrange the plots in grids
gridExtra::grid.arrange(gn1,gn2,gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,gn7,gn8,ncol=2)
gridExtra::grid.arrange(gn9,gn10,ncol=2)


# #Replace all outliers with NA and impute
for(i in cnames){
  val = emp_abs[,i][emp_abs[,i] %in% boxplot.stats(emp_abs[,i])$out]
  print(paste(i,length(val)))
  emp_abs[,i][emp_abs[,i] %in% val] = NA
}

# Outlier-missing value calculation
# Get number of missing values after replacing outliers as NA
outlier_missing_values = data.frame(sapply(emp_abs,function(x){sum(is.na(x))}))
outlier_missing_values$Columns = row.names(outlier_missing_values)
row.names(outlier_missing_values) = NULL
names(outlier_missing_values)[1] = "Out_Miss_perc"
outlier_missing_values$Out_Miss_perc = ((outlier_missing_values$Out_Miss_perc/nrow(emp_abs))*100)
outlier_missing_values = outlier_missing_values[,c(2,1)]
outlier_missing_values = outlier_missing_values[order(-outlier_missing_values$Out_Miss_perc),]
outlier_missing_values

# Visualization of missing value
ggplot(data = outlier_missing_values[1:6,], aes(x=reorder(Columns, -Out_Miss_perc),y = Out_Miss_perc))+
  geom_bar(stat = "identity",fill = "green")+xlab("Parameter")+
  ggtitle("Outlier Missing data percentage (Train)") + theme_bw()

# KNN imputation
emp_abs = knnImputation(emp_abs, k=3)
sum(is.na(emp_abs))

# Data copy
df = emp_abs
# emps_abs = df
################################## Feature Selection ###########################

## Correlation Plot
# corrgram library helps to plot correlation plot
corrgram(emp_abs[,numeric_index], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

#Check for multicollinearity using VIF
vifcor(numeric_data)

## Dimension Reduction
emp_abs = subset.data.frame(emp_abs,select = -c(Weight))

numeric_data = subset.data.frame(numeric_data,select = -c(Weight))
################################# Feature scaling ##############################
# Norma Q-Q plot 
qqnorm(emp_abs$Transportation.expense)
# Normality checks with histogram
h1<-ggplot(emp_abs, aes_string(emp_abs$Transportation.expense))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Transportation.expense")
h2<-ggplot(emp_abs,aes_string(emp_abs$Distance.from.Residence.to.Work))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Distance.from.Residence")
h3<-ggplot(emp_abs,aes_string(emp_abs$Service.time))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Service.time")
h4<-ggplot(emp_abs,aes_string(emp_abs$Age))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Age")
h5<-ggplot(emp_abs,aes_string(emp_abs$Work.load.Average.day.))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Work.load.Average.day.")
h6<-ggplot(emp_abs,aes_string(emp_abs$Hit.target))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Hit.target")
h7<-ggplot(emp_abs,aes_string(emp_abs$Height))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Height")
h8<-ggplot(emp_abs,aes_string(emp_abs$Body.mass.index))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Body.mass.index")
h9<-ggplot(emp_abs,aes_string(emp_abs$Absenteeism.time.in.hours))+
  geom_histogram(fill="DarkSlateBlue",colour ='black',bins = 30)+ggtitle("Absenteeism.time.in.hours")

gridExtra::grid.arrange(h1,h2,h3,ncol=1)
gridExtra::grid.arrange(h4,h5,h6,ncol=1)
gridExtra::grid.arrange(h7,h8,h9,ncol=1)

# Remove dependent variable
numeric_index = sapply(emp_abs,is.numeric) # Selecting only numeric
numeric_data = emp_abs[,numeric_index]
numeric_col = names(numeric_data)
numeric_col = numeric_col[-9]

for (i in numeric_col) {
  print(i)
  emp_abs[,i] = (emp_abs[,i] - min(emp_abs[,i]))/
    (max(emp_abs[,i]-min(emp_abs[,i])))
  
}
################################## Sampling #####################################
# Creating dummy variables for categorical variables
categorical_names = names(categorical_data)
emp_abs = dummy.data.frame(emp_abs,categorical_names)

#Splitting data into train and test data
set.seed(1)
train_index = sample(1:nrow(emp_abs), 0.8*nrow(df))        
train = emp_abs[train_index,]
test = emp_abs[-train_index,]

#Principal component analysis
p_c = prcomp(train)
#Standard deviation for principal component
pr_sd = p_c$sdev
#Variance calculation
pr_v = pr_sd^2
#Proportion of variance explained
pr_variance = pr_v/sum(pr_v)

# 97+ % data variance explained by 40 components
plot(cumsum(pr_variance), xlab = "Principal components",ylab = "cumulative proportion of variance",
     type = "b",main = "PCA")

# Adding a training set with principal components
train_1 = data.frame(Absenteeism.time.in.hours = train$Absenteeism.time.in.hours,p_c$x)
#Transforming data
train_1 = train_1[,1:40]

test_1 = predict(p_c,newdata = test)
test_1 = as.data.frame(test_1)
test_1 = test_1[,1:40]
# ###################################### Decision tree model ###################################
#rpart for regression(method = 'anova' for regression model, method = 'class' for classification model)
fit = rpart(Absenteeism.time.in.hours~., data = train_1, method = "anova")
# 
# Predict for new test cases
predictions = predict(fit,test_1)
#
dt_pred = data.frame("actual"= test[,115],"DT_model" = predictions)
head(dt_pred)

#Calculate MAE, RMSE, R-squared for testing data 
print(postResample(pred = predictions, obs = test$Absenteeism.time.in.hours))

################################ Random Forest model #######################################
###Random Forest (importance = TRUE , it shows important trees in the model)
RF_model = randomForest(Absenteeism.time.in.hours ~ ., train_1, importance = TRUE, ntree = 500)

#Predict test data using random forest model
RF_Predictions = predict(RF_model, test_1)

#Create data frame for actual and predicted values
dt_pred_RF = data.frame("actual"= test[,115],"RF_model" = RF_Predictions)
head(dt_pred_RF)

#Calculate MAE, RMSE, R-squared for testing data 
print(postResample(pred = RF_Predictions, obs = test$Absenteeism.time.in.hours))
################################# Linear regression model #################################
#Train the model using training data
LR_model = lm(Absenteeism.time.in.hours ~ ., data = train_1)

#Get the summary of the model
summary(LR_model)

#Predict the test cases
LR_predictions = predict(LR_model,test_1)

#Create data frame for actual and predicted values
dt_pred_LR = data.frame("actual"= test[,115],"LR_model" = LR_predictions)
head(dt_pred_LR)

#Calculate MAE, RMSE, R-squared for testing data 
print(postResample(pred = LR_predictions, obs = test$Absenteeism.time.in.hours))

# Save the model implemented document
write.csv(emp_abs,"emp_abs_dummy_m_R.csv",row.names = F)
