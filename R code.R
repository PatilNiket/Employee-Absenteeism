rm(list=ls())

setwd("C:/Users/User/Desktop/Project 2/File")
getwd()

#Load Libraries
x = c("xlsx","ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "dummies", "e1071", "Information",
      "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees')

#install.packages(x)
lapply(x, require, character.only = TRUE)


## Read the data
df_Absenteeism = read.xlsx('Absenteeism.xlsx', sheetIndex = 1)

#---------------------------Exploratory Data Analysis--------------------------------------#

head(df_Absenteeism)
dim(df_Absenteeism)
names(data)
summary(df_Absenteeism)
colnames(df_Absenteeism)

#------------------------------------Missing Values Analysis---------------------------------------------------#
#calculating missing value 

missing_val = data.frame(apply(df_Absenteeism,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"

#Calculating percentage missing value
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(df_Absenteeism)) * 100

# Sorting missing_val in Descending order
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL

# Reordering columns
missing_val = missing_val[,c(2,1)]

# # Missing value plot 
ggplot(data = missing_val[1:100,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
geom_bar(stat = "identity",fill = "red")+xlab("Variables")+
ggtitle("Missing data percentage") + theme_bw()

# Let us consider which method is sutable for our data to remove missing value
df_Absenteeism$Distance.from.Residence.to.Work[85]

# Actual Value = 31
# Mean = 29.66576
# Median = 26
# KNN = 31


#Mean Method
#df_Absenteeism$Distance.from.Residence.to.Work[85]=NA
#df_Absenteeism$Distance.from.Residence.to.Work[is.na(df_Absenteeism$Distance.from.Residence.to.Work)] = mean(df_Absenteeism$Distance.from.Residence.to.Work, na.rm = T)

#Median Method
#df_Absenteeism$Distance.from.Residence.to.Work[85]=NA
#df_Absenteeism$Distance.from.Residence.to.Work[is.na(df_Absenteeism$Distance.from.Residence.to.Work)] = median(df_Absenteeism$Distance.from.Residence.to.Work, na.rm = T)

# kNN Imputation
df_Absenteeism$Distance.from.Residence.to.Work[85]=NA
df_Absenteeism = knnImputation(df_Absenteeism, k = 3)

# Checking for missing value
sum(is.na(df_Absenteeism))

##################################### Deffrenciating data #######################################

cnames = c('Transportation.expense','Distance.from.Residence.to.Work', 'Service.time', 'Age',
            'Work.load.Average', 'Hit.target', 'Weight', 'Height', 'Body.mass.index', 'Absenteeism.time.in.hours')

cat_names= c('ID','Reason.for.absence','Month.of.absence','Day.of.the.week',
                     'Seasons','Disciplinary.failure', 'Education', 'Son', 'Social.drinker',
                     'Social.smoker',  'Pet')


#-------------------------------------Outlier Analysis-------------------------------------#

# Boxplot for continuous variables

for(i in 1:length(cnames))
  {assign(paste0("PK",i),ggplot(aes_string(y=(cnames[i]),x="Absenteeism.time.in.hours"),d=subset(df_Absenteeism))
         +geom_boxplot(outlier.colour = "Red",outlier.shape = 18,outlier.size = 1, fill="skyblue4")+theme_gray()
         +stat_boxplot(geom = "errorbar", width=0.5) +labs(y=cnames[i],x="Absenteeism time(Hours)")
         +ggtitle("Box Plot of Absenteeism for",cnames[i]))

}

for (i in 1:length(cnames))

# ## Plotting plots together
gridExtra::grid.arrange(PK1,PK2,PK3,ncol=3)
#gridExtra::grid.arrange(PK4,PK6,PK7,ncol=3)
#gridExtra::grid.arrange(PK8,PK9,PK10,ncol=3)

# #Remove outliers using boxplot method

# #loop to remove from all variables
for(i in cnames)
{
  print(i)
  val = df_Absenteeism[,i][df_Absenteeism[,i] %in% boxplot.stats(df_Absenteeism[,i])$out]
  #print(length(val))
  df_Absenteeism = df_Absenteeism[which(!df_Absenteeism[,i] %in% val),]
}
dim(df_Absenteeism)


#--------------------------------Feature Selection--------------------------------------------#

#Correlation Analysis for continuous variables-

corrgram(df_Absenteeism[,cnames],order=FALSE,upper.panel = panel.pie,
         text.panel = panel.txt,font.labels =1, main="Correlation plot")


# Weight is highly corrlrted

#Anova Test for categorical variable-


for(i in cat_names){
  print(i)
  Anova_result= summary(aov(formula = Absenteeism.time.in.hours~df_Absenteeism[,i] ,df_Absenteeism))
  print(Anova_result)
}

# Dimension Reduction
df_Absenteeism = subset(df_Absenteeism, select = -c(Weight, Month.of.absence, Seasons , Education , Social.smoker , Pet))

dim(df_Absenteeism)

################################ FUTURE SCALING ################################

hist(df_Absenteeism$Transportation.expense,col="Red",main="Histogram of Transportation Expense")

# Updating the continuous and catagorical variable
cnames = c('Distance.from.Residence.to.Work', 'Service.time', 'Age','Work.load.Average', 
           'Transportation.expense','Hit.target', 'Height', 'Body.mass.index')

cat_names = c('ID','Reason.for.absence','Disciplinary.failure', 'Social.drinker', 'Son', 'Day.of.the.week')

# Normalization
for(i in cnames)
{
  print(i)
  df_Absenteeism[,i] = (df_Absenteeism[,i] - min(df_Absenteeism[,i]))/(max(df_Absenteeism[,i])-min(df_Absenteeism[,i]))
}

df = df_Absenteeism
#df_Absenteeism=df

# save pre process file 
write.csv(df_Absenteeism,"Absenteeism_Pre_proc.csv",row.names=FALSE)

######################### dummy variables for categorical variables #######################
library(mlr)
df_Absenteeism = dummy.data.frame(df_Absenteeism, cat_names)

#=============================########## Model Devlopment ##########=================================================#
#Clean the Environment-
rmExcept("df_Absenteeism")

##################################################DEVIDE DATA INTO 80:20 #####################################################

#Divide data into train and test
set.seed(123)
train.index = sample(1:nrow(df_Absenteeism), 0.8 * nrow(df_Absenteeism))                                                                                                                                                                          
train = df_Absenteeism[ train.index,]
test  = df_Absenteeism[-train.index,]


#----------------------Dimensionality Reduction using PCA-------------------------------#


prin_comp = prcomp(train) #PCA

std_dev = prin_comp$sdev  #compute SD of each principal component

pr_var = std_dev^2 # variance calculation

prop_varex = pr_var/sum(pr_var) ##proportion of variance

#cumulative plot
plot(cumsum(prop_varex), xlab = "Principal Component", ylab = "Cumulative Proportion of Variance", 
      type = "b")

#add a training set with principal components
train.data = data.frame(Absenteeism.time.in.hours = train$Absenteeism.time.in.hours, prin_comp$x)

# From the above plot selecting 30 components since it explains almost 95+ % data variance
train.data =train.data[,1:30]

#transform test into PCA
test.data = predict(prin_comp, newdata = test)
test.data = as.data.frame(test.data)

#select the first 30 components
test.data=test.data[,1:30]


################################### DT for Regression ##########################################

                              ----------#DT Model#-----------

library(rpart)
fit_DT = rpart(Absenteeism.time.in.hours ~., data = train.data, method = "anova")

#Summary of DT model
#summary(fit_DT)

#Lets predict for training data
DT_train = predict(fit_DT, train.data)

#Lets predict for training data
DT_test = predict(fit_DT,test.data)

#######################################Error metrics to calculation###########################

print(postResample(pred = DT_train, obs = train$Absenteeism.time.in.hours)) # For train

print(postResample(pred = DT_test, obs = test$Absenteeism.time.in.hours)) # For Test 

############################### Visulaization to check the model performance ###########################

           ########### TRAIN ##############
plot(train$Absenteeism.time.in.hours,type="l",lty=1.8,col="Red")
lines(DT_train,type="l",col="blue")

          ########### TEST ###############
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green")
lines(DT_test,type="l",col="Pink")


####################################  RF  ###################################################

set.seed(140)
library(randomForest)  #Library for randomforest 
library(inTrees)       #Library for intree transformation

#Develop Model on training data
fit_RF = randomForest(Absenteeism.time.in.hours~., data = train.data)

#Lets predict for training data
RF_train = predict(fit_RF, train.data)

RF_test = predict(fit_RF,test.data)

#######################################Error metrics to calculation###########################

print(postResample(pred = RF_train, obs = train$Absenteeism.time.in.hours)) # For Train

print(postResample(pred = RF_test, obs = test$Absenteeism.time.in.hours)) # For Test 

############################### Visulaization to check the model performance ###########################

########### TRAIN ##############
plot(train$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green")
lines(RF_train,type="l",col="Blue")

########### TEST ###############
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Red")
lines(RF_test,type="l",col="Blue")


##########################################Linear Regression############################################

set.seed(140)

#Develop Model on training data
fit_LR = lm(Absenteeism.time.in.hours ~ ., data = train.data)

#Lets predict for training data
LR_train = predict(fit_LR, train.data)

#Lets predict for testing data
LR_test = predict(fit_LR,test.data)

----------------------------# Error metrics to calculation #-----------------------------------------

print(postResample(pred = LR_train, obs = train$Absenteeism.time.in.hours)) # For Train

print(postResample(pred = LR_test, obs = test$Absenteeism.time.in.hours)) # For Test 

############################### Visulaization to check the model performance ###########################

########### TRAIN ##############
plot(train$Absenteeism.time.in.hours,type="l",lty=1.8,col="red")
lines(LR_train,type="l",col="blue")

########### TEST ###############
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green")
lines(LR_test,type="l",col="Blue")


########################################### XGBoost ########################################

set.seed(140)

#Develop Model on training data
fit_XGB = gbm(Absenteeism.time.in.hours~., data = train.data, n.trees = 500, interaction.depth = 2)

#Lets predict for training data
XGB_train = predict(fit_XGB, train.data, n.trees = 500)

#Lets predict for testing data
XGB_test = predict(fit_XGB,test.data, n.trees = 500)

----------------------------# Error metrics to calculation #-----------------------------------------

print(postResample(pred = XGB_train, obs = train$Absenteeism.time.in.hours)) # For train
print(postResample(pred = XGB_test, obs = test$Absenteeism.time.in.hours)) # For Test 

############################### Visulaization to check the model performance ###########################

########### TRAIN ##############
plot(train$Absenteeism.time.in.hours,type="l",lty=1.8,col="Red")
lines(XGB_train,type="l",col="Blue")

########### TEST ###############
plot(test$Absenteeism.time.in.hours,type="l",lty=1.8,col="Green")
lines(XGB_test,type="l",col="Blue")








