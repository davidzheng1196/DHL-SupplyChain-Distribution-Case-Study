rm(list=ls())
library(readxl)
library(xlsx)
library(dplyr)
library(tidyr)


shipments<-read_excel('Worksheet in Distribution Case Study Problem (3).xls',sheet=1)
stores<-read_excel('Worksheet in Distribution Case Study Problem (3).xls',sheet=2)
shipments<-as.data.frame(shipments)
stores<-as.data.frame(stores)

###############################
###Analyzing Demand############
###############################
groupby_date<- shipments %>%group_by(`Shipment Date`)

analysis<-groupby_date%>%summarise(
  weight_groupby<- mean(`Weight (Lbs)`),
  cases_groupby<-mean(`Cases`),
  space_groupby<-sum(`Cubic Feet`),
  cases_sum_groupby<-sum(Cases),
  pallet_sum_groupby<-sum(ceiling(Cases/45))
)
plot(ts(analysis[c(2,3)]),main='Demand over Time')
## The demand peaked at Monday and Friday
summary(shipments['Weight (Lbs)'])
summary(shipments['Cases'])
sqrt(var(shipments['Weight (Lbs)'])) #Standard Deviation of every single Weight (Lbs)
sqrt(var(shipments['Cases']))        #Standard Deviation of every single Cases



##Using BootStrap Sampling Method to get average daily demand
weight_list<-c()
for (i in 1:60){
  bootstrap<-c()
  for (j in 1:1000){
    bootstrap[j]<-mean(sample(shipments[shipments['Customer Number']==99+i,]['Weight (Lbs)'][,1],28,replace=TRUE))
  }
  weight_list[i]<-mean(bootstrap)
}

stores<-cbind(stores,weight_list) ##creating new column


###############################
###Center of Gravity method####
###############################
#Cx= sum (dix * Vi)/ sum (Vi)
#Cy= sum (diy * Vi)/ sum (Vi)
#We calculate the X and Y coordinates using these equations 
#where Cx is the X (horizontal axis) coordinate for the new facility. 
#Cy is the Y (vertical axis) coordinate for the new facility, 
#dix is the X coordinate of the existing location, 
#diy is the Y coordinate of the existing location, 
#and Vi is the volume of goods moved to or from the ith location.
Cx1<-c()
Cx2<-c()
Cy1<-c()
Cy2<-c()
for (i in 1:nrow(stores)){
  Cx1[i]<-stores[i,6]*stores[i,7]
  Cx2[i]<-stores[i,7]
  Cy1[i]<-stores[i,5]*stores[i,7]
  Cy2[i]<-stores[i,7]
}
Cx<-sum(Cx1)/sum(Cx2)
Cy<-sum(Cy1)/sum(Cy2)
(DC_coordinate<-c(Cy,Cx))


###############################
###Travel time matrices########
###############################
library(osrm)
osrm_list<-stores[,c('Customer Number','X','Y')]
osrm_list<- rbind(osrm_list,c(160,-84.3,34))  #The ID for WH is 160
tt_matrices<-osrmTable(loc=osrm_list[1:61,c('Customer Number','X','Y')])
durations_matrix<-tt_matrices$durations


###############################
###Travel distance matrices####
###############################
distance_matrix<-durations_matrix
for (row in 1:nrow(distance_matrix)){
  for (column in 1:ncol(distance_matrix)){
    distance_matrix[row,column]<-(durations_matrix[row,column]+durations_matrix[column,row])/2/60*55
  }
}

write.xlsx(distance_matrix, "distance_matrix.xlsx",col.names = TRUE,row.names = TRUE)

#Go to Jupyter Note Book
######################################################################################


###############################
###Display the Network#########
###############################
library(geosphere)
library(maps)

path<-c(60, 41, 28, 29, 9, 17, 30, 6, 46, 13, 31, 21, 27, 55, 
        45, 35, 1, 14, 47, 23, 37, 2, 57, 15, 8, 22, 49, 7, 
        54, 36, 34, 42, 58, 43, 33, 11, 53, 19, 4, 44, 59, 16, 
        3, 40, 56, 18, 51, 20, 5, 25, 52, 38, 48, 26, 0, 24, 
        10, 50, 32, 12, 39, 60)
for (i in 1:length(path)){
  path[i]<-path[i]+1
}

long<-c()
lat<-c()
for (i in 1:length(path)){
  long[i]<-osrm_list[path[i],2]
  lat[i]<-osrm_list[path[i],3]
}

map('state','georgia',xlim=range(long),ylim=range(lat))
points(x=long,y=lat,col='red',cex=1,pch=1)
lines(long,lat,pch=10)
points(-84.3,34,col='blue',cex=3,pch=7)

map('county','georgia',xlim=range(long),ylim=range(lat))
points(x=long,y=lat,col='red',cex=1,pch=1)
lines(long,lat,pch=10)
points(-84.3,34,col='blue',cex=3,pch=7)
##Keep in mind that the lines in this network model only indicates 
##the sequences of TSP(Traveling Sales Person) using Christofides Algorithm. 


###############################
###Assign TruckLoad ###########
###############################
TL_path<-path[2:61]
TL<- data.frame(matrix(NA,nrow=length(unique(shipments$`Shipment Date`)),ncol=length(TL_path)))
rownames(TL)<-unique(shipments$`Shipment Date`)
colnames(TL)<-TL_path

#############
max_weight<-45000    ##maximum lbs carrying limit for 53 foot trailer
max_capacity<-26     ##maximum number of pallets for 53 foot trailer
max_time<-540        ##Deliver between 08:00 - 17:00
#############

colnames(durations_matrix)<-1:61
rownames(durations_matrix)<-1:61

layer<-0
for (row in 1:nrow(TL)){
  trailer<-1
  weight<-0
  capacity<-0
  time<-0
  for (col in 1:ncol(TL)){
    weight<-weight+shipments[path[col+1]+layer,7]
    capacity<-capacity+ceiling(shipments[path[col+1]+layer,9]/45)
    time<-time+durations_matrix[path[col],path[col+1]]+30
    
    if((weight< max_weight) && (capacity< max_capacity) && (time+durations_matrix[path[col+1],61]< max_time)){
      TL[row,col]<-trailer
    }else{
      trailer<-trailer+1
      TL[row,col]<-trailer
      weight<-0
      capacity<-0
      time<-0
      
      weight<-weight+shipments[path[col+1]+layer,7]
      capacity<-capacity+ceiling(shipments[path[col+1]+layer,9]/45)
      time<-time+durations_matrix[61,path[col+1]]+30
    }
  }
  layer<-layer+60
}

colnames(TL)<-as.numeric(names(TL))+99
write.xlsx(TL, "Distribution Design.xlsx",col.names = TRUE,row.names = TRUE,sheetName='Route_Plan') ## output the Truckload plan as excel file


###############################
###Drive/ Work Time  ##########
###############################
##Drive Time
number_of_equipment<-c()
for (row in 1:nrow(TL)){
  for (col in 1:ncol(TL)){
    number_of_equipment<-rbind(number_of_equipment,TL[row,col])
  }
}
number_of_equipment<-unique(number_of_equipment)

drive_time<-as.data.frame(matrix(NA,nrow=length(number_of_equipment),ncol=nrow(TL)))
colnames(drive_time)<-rownames(TL)

for(col in 1:ncol(drive_time)){
  for(row in 1:nrow(drive_time)){
    k<-number_of_equipment[row]
    if(k %in% TL[rownames(TL)==names(drive_time[col]),]){
      single_route<-c(61,as.numeric(colnames(TL[col,TL[col,]==k]))-99,61)
      time_spent<-0
      for (i in 1:(length(single_route)-1)){
        time_spent<-time_spent+durations_matrix[single_route[i+1],single_route[i]]
      }
      drive_time[row,col]<-time_spent
    }else{
      drive_time[row,col]<-NA
    }
  }
}
drive_time<-t(drive_time)


##Work Time
work_time<-as.data.frame(matrix(NA,nrow=length(number_of_equipment),ncol=nrow(TL)))
colnames(work_time)<-rownames(TL)

for(col in 1:ncol(work_time)){
  for(row in 1:nrow(work_time)){
    k<-number_of_equipment[row]
    if(k %in% TL[rownames(TL)==names(work_time[col]),]){
      single_route<-c(61,as.numeric(colnames(TL[col,TL[col,]==k]))-99,61)
      time_spent<-0
      for (i in 1:(length(single_route)-1)){
        time_spent<-time_spent+durations_matrix[single_route[i],single_route[i+1]]+30
      }
      work_time[row,col]<-time_spent-30
    }else{
      work_time[row,col]<-NA
    }
  }
}
work_time<-t(work_time)


###############################
###Driving Distance############
###############################
driving_distance<-as.data.frame(matrix(NA,nrow=length(number_of_equipment),ncol=nrow(TL)))
colnames(driving_distance)<-rownames(TL)
colnames(distance_matrix)<-1:61
rownames(distance_matrix)<-1:61

for(col in 1:ncol(driving_distance)){
  for(row in 1:nrow(driving_distance)){
    k<-number_of_equipment[row]
    if(k %in% TL[rownames(TL)==names(driving_distance[col]),]){
      single_route<-c(61,as.numeric(colnames(TL[col,TL[col,]==k]))-99,61)
      time_spent<-0
      for (i in 1:(length(single_route)-1)){
        time_spent<-time_spent+distance_matrix[single_route[i+1],single_route[i]]
      }
      driving_distance[row,col]<-round(time_spent,2)
    }else{
      driving_distance[row,col]<-NA
    }
  }
}
driving_distance<-t(driving_distance)

###############################
###Fuel Cost###################
###############################
mpg<-6.11        ##According to "The State of Fuel Economy in Trucking"
gas_price<-2.38  ##Quoted from Georgia 2019 average price of gas

fuel_cost<-driving_distance
for (row in 1:nrow(fuel_cost)){
  for (col in 1:ncol(fuel_cost)){
    fuel_cost[row,col]<-round(fuel_cost[row,col]/mpg*gas_price,2)
  }
}


###############################
###Trucking Cost###############
###############################
marginal_cost_per_mile<-1.553  ##According to "ATRI Analysis of the Operational Costs of Trucking"

trucking_cost<-driving_distance
for (row in 1:nrow(trucking_cost)){
  for (col in 1:ncol(trucking_cost)){
    trucking_cost[row,col]<-round(trucking_cost[row,col]*marginal_cost_per_mile,2)
  }
}


###############################
###Direct Labor(Head Count)####
###############################
space_required<-analysis[4]
rownames(space_required)<-unique(shipments$`Shipment Date`)
inbound_cases<-ceiling(analysis[5])
rownames(inbound_cases)<-unique(shipments$`Shipment Date`)

######################Set Labor Productivity
unloading<-500          #500 cases per hour per capita
loading<-40             #40 pallets per hour per capita
sorting<-150            #150 cases per hour per capita
shift_hour<-8           #8 hour night shift
utilization_rate<-0.80  #80% utilization rate
#######################
headcount_unloading<-ceiling(inbound_cases/(unloading*shift_hour*utilization_rate))
headcount_sorting<-ceiling(inbound_cases/(sorting*shift_hour*utilization_rate))
headcount_loading<-ceiling(analysis[6]/(loading*shift_hour*utilization_rate))

Total_Head_count<-cbind(headcount_loading,headcount_sorting,headcount_unloading)
Total_Head_count<-cbind(Total_Head_count,rowSums(Total_Head_count))
colnames(Total_Head_count)<-c('Unloading','Sorting','Loading','Sum')


###############################
###Cost Report#################
###############################
hourly_wage<-15      # $15 hourly wage

worker_salary<-Total_Head_count[,1:3]*hourly_wage*shift_hour
Total_Cost<-cbind(worker_salary,rowSums(trucking_cost,na.rm=TRUE))
Total_Cost<-cbind(Total_Cost,rowSums(Total_Cost))
colnames(Total_Cost)<-c('Unloading','Sorting','Loading','Trucking','Sum')



###################################################
##Load results into Excel sheet
write.xlsx(drive_time, "Distribution Design.xlsx",append=TRUE,col.names = TRUE,row.names = TRUE,showNA=FALSE,sheetName='Drive_Time')

write.xlsx(work_time, "Distribution Design.xlsx",append=TRUE,col.names = TRUE,row.names = TRUE,showNA=FALSE,sheetName='Work_Time')

driving_distance<-cbind(driving_distance,rowSums(driving_distance,na.rm=TRUE))
colnames(driving_distance)<-c(1:(ncol(driving_distance)-1),'SUM')
write.xlsx(driving_distance, "Distribution Design.xlsx",append=TRUE,col.names = TRUE,row.names = TRUE,showNA=FALSE,sheetName='Driving_Distance')

fuel_cost<-cbind(fuel_cost,rowSums(fuel_cost,na.rm=TRUE))
colnames(fuel_cost)<-c(1:(ncol(fuel_cost)-1),'SUM')
write.xlsx(fuel_cost, "Distribution Design.xlsx",append=TRUE,col.names = TRUE,row.names = TRUE,showNA=FALSE,sheetName='Fuel_Cost')

trucking_cost<-cbind(trucking_cost,rowSums(trucking_cost,na.rm=TRUE))
colnames(trucking_cost)<-c(1:(ncol(trucking_cost)-1),'SUM')
write.xlsx(trucking_cost, "Distribution Design.xlsx",append=TRUE,col.names = TRUE,row.names = TRUE,showNA=FALSE,sheetName='Trucking_Cost')

write.xlsx(Total_Cost, "Distribution Design.xlsx",append=TRUE,col.names = TRUE,row.names = TRUE,showNA=FALSE,sheetName='Total_Cost')
