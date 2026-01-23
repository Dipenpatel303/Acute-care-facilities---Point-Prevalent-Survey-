***********************************************************************************
***********************************************************************************
                                             Tennessee Department of Health
             CEDEP / Healthcare Associated Infections and Antimicrobial Stewardship Program
***********************************************************************************
***********************************************************************************
       Project: Acute Care Hospitals Point Prevalence Survey    

Programmer: Yusuf 

           Date: September 2018   

Modified: Dipen M Patel 

       Version:2.0

***********************************************************************************
***********************************************************************************
Updates :

***********************************************************************************
***********************************************************************************
Notes:  1- Process the original RedCap SAS file separately before calling it with the %include command                                                                       *
           2- Updates are required where 'UPDATE' is mentioned  -- Do "Ctrl+F" to search for these fields --In upper case

***********************************************************************************
***********************************************************************************;


*Setting system options for macro debugging purpose;              
options  symbolgen mprint mlogic mcompilenote=all;

/* Setting global Macro variables for quarter and output path */

 *Automatic setting of Quarter and Year Macro variables;     
%macro time ; 
%global currentq;
%global currenty;

data _null_;
call symputx ('currentq', qtr(intnx('qtr',today(),-1)));
run;

%if &currentq=4 %then %do;

data _null_;
call symputx ('currenty', year (today())-1);
run;
                                          %end;

                                  %else %do;

data _null_;
call symputx ('currenty', year (today()));
run;

                                           %end;
%mend;

%time;
%put &currentq &currenty;

*Automatic setting of start and end date;
data _null_;
x=intnx ('day',today(),-60);
call symputx ('enddate', put(intnx('month',x, 0, 'E'), date9.));
call symputx ('startdate', put(intnx ('qtr',x,-5, 'B'), date9.));
run;
/*%put &startdate &enddate;*/

*Other macro variables;
%let quarter=Q&currentq-&currenty;
%let path=%str(H:\NHSN\Antimicrobial Stewardship\Antibiotic Use Survey Data\ACF AU Survey\&currenty Q&currentq Analysis\Packets);*Path=where the final PDF packages will be saved;
%let inpath=%str(H:\NHSN\Antimicrobial Stewardship\Antibiotic Use Survey Data\ACF AU Survey\&currenty Q&currentq Analysis\SAS programs);*Inpath=the path where the original Redcap survey SAS code is saved;
%put &inpath;
/* Importing Quarter data and cleaning */

%let file=AntimicrobialUseSurv_SAS_2024-02-23_1057;*UPDATE file name;
%include "H:\NHSN\Antimicrobial Stewardship\Antibiotic Use Survey Data\ACF AU Survey\2023 Q4 Analysis\SAS programs\AntimicrobialUseSurv_SAS_2024-02-23_1057.sas";

/*  REDCAP Data cleaning &  manipulation */

   
proc format;*Format to be applied to the quarter variable;
picture myfmt
   low-high='%Y-Q%q ' (datatype=date);
   run;

data redcap1(label='Past6_quarters')  redcap2(label='Current_quarter'); 
 set redcap;
 *Deleting incomplete surveys;
 if survey_of_antibiotic_v_1 eq 2 ;
 *Correcting Vandy 2020 Q3 submission date;
 if record_id=3106 then surveydate='01jul2020'd;
 *Correcting Blount Q2-2021 entry;
 if record_id=3243 then surveydate='30jun2021'd;
 * Defining the desired quarter data: Past six quarters;
if "&startdate"d<=surveydate<="&enddate"d; 
quarter=put (surveydate, myfmt.);
*Changing hospital names to match with Bed Size data later;
 hospname1= upcase(hospname);
  if index(hospname1,'HUNTINGDON')  then hospname='Baptist Memorial Hospital - Huntingdon';
  else if index(hospname1,'CREST') or index(hospname1,'CRREST') then hospname='NorthCrest Medical Center';
  else if index (hospname1,'UNION CITY') or record_id=3293  then hospname='Baptist Memorial Hospital - Union City'; 
  else if index (hospname1,'SKYRIDGE')  then hospname='Tennova Healthcare - Cleveland';
  else if index (hospname1,'BOLIVAR')  then hospname='Bolivar General Hospital';
  else if index(hospname1, 'EAST TENNESSEE CHILDREN')  then hospname="East Tennessee Children Hospital";
 else if index (hospname1,'HOLSTON')  then hospname='Holston Valley Medical Center';
 else if index (hospname1,'MEMPHIS')  then hospname='Baptist Memorial Hospital - Memphis';
  else if index (hospname1,'TAKOMA')  then hospname='Takoma Regional Hospital';
 else if index (hospname1,'MIDTOWN')  then hospname='St. Thomas Midtown Hospital';
 else if index (hospname1,'SUMNER')  then hospname='Sumner Regional Medical Center';
 else if index (hospname1,'VANDERBILT')  then hospname='Vanderbilt Medical Center';
 else if index (hospname1,'PARKWEST')  then hospname='Parkwest Medical Center- Knoxville';
 else if index (hospname1,'MADISON') or index (hospname1,'MADISION')  then hospname='Jackson Madison County General Hospital';
 else if index (hospname1,'CENTENNIAL') or index (hospname1,'CENTENIAL') then hospname='Centennial Medical Center';
 else if index (hospname1,'FOLLETTE') ge 1 then hospname='Tennova Healthcare - Lafollette Med Ctr';
 else if index (hospname1,'REGIONAL HOSPITAL OF JACKSON') or index(hospname1,'TENNOVA REGIONAL') then hospname='Tennova Healthcare - Regional Jackson';
 else if index (hospname1, 'BRISTOL') ge 1 then hospname='Bristol Regional Medical Center';
  else if index (hospname1, 'CHI MEMORIAL') ge 1 then hospname='CHI Memorial';
  else if index (hospname1, 'MAURY') ge 1 then hospname='Maury Regional Medical Center';
  else if index (hospname1, 'LECONTE') ge 1 then hospname='LeConte Medical Center';
else  if index (hospname1, 'MILAN') or index (hospname1, 'WTH')  then hospname='West Tennessee Healthcare Milan Hospital';
 else if index (hospname1, 'CAMDEN') ge 1 then hospname='West Tennessee Healthcare Camden Hospital';
 else if index (hospname1, 'UT MEDICAL') or index (hospname1, 'UNIVERSITY OF TENNESSEE')  then hospname='University of Tennessee Medical Ctr';
 else if index (datacollector,'Giles') or index (datacollector,'Shelley') or index (hospname1, 'STARR')  then hospname='Starr Regional Medical Center - Athens';
 else if index (hospname1, 'LOUDON') or index (hospname1, 'LOUDOUN')  then hospname='Fort Loudoun Medical Center';
 else if index (hospname1, 'CURAHEALTH') ge 1 then hospname='Curahealth Nashville';
 else if index (hospname1, 'HORIZON') ge 1 then hospname='Horizon Medical Center';
 else if index (hospname1, 'NORTH KNOXVILLE') ge 1 then hospname='Tennova Healthcare - North Knoxville Med Ctr';
 else if index (hospname1, 'DELTA') ge 1 then hospname='Delta Medical Center';
  else if index (hospname1, 'FRANCIS') ge 1 then hospname='St. Francis Bartlett';
  else if index (hospname1, 'WAYNE') ge 1 then hospname='Wayne Medical Center';
   else if index(hospname1, 'GREENEVILLE COMMUNITY HOSPITAL WEST')  then hospname="Greeneville Community Hospital";
  else if index(hospname1,'GREENEVILLE COMMUNITY HOSPITAL EAST')  then hospname="Greeneville Community Hospital";
  else if index(hospname1,'GREENEVILLE COMMUNITY HOSPITAL')  then hospname="Greeneville Community Hospital";

else hospname=hospname;
  if index(hospname1,'BLOUNT')  then hospname='Blount Memorial Hospital';
  if hospname1='HOSPITAL' then hospname='Starr Regional Medical Center - Athens';
  if hospname1='NORTHCERST' or index(hospname1,'NORTHCEST') then hospname='NorthCrest Medical Center';
  if index(hospname1,'MARTIN') or index(hospname1,'VOLUNTEER') then hospname='West Tennessee Healthcare Volunteer Hospital-Martin';
  *converting NA values to missing;
if upcase(otherbetalactam) in ('NA', 'N/A') then otherbetalactam=' ';
if upcase(nonbetalactam) in ('NA', 'N/A') then nonbetalactam=' ';
*Using Array to convert character variables to numeric;
array charvar[*]  $ anyabx vanco linezolid daptomycin cefotax ceftaz piptazo meropenem levomoxi cipro tigecycline otherbetalactam nonbetalactam ;
array numvar [*]  anyabx1 vanco1 linezolid1 daptomycin1 cefotax1 ceftaz1 piptazo1 meropenem1 levomoxi1 cipro1 tigecycline1 otherbetalactam1 nonbetalactam1;
do i=1 to dim(charvar);
numvar[i]=input (charvar[i], 4.);
if charvar[i] eq ' '  then numvar[i]=0;
end;
*Renaming the numeric variables created above under the ARRAY statement;
rename anyabx1=anyabx vanco1=vanco linezolid1=linezolid daptomycin1=daptomycin cefotax1=cefotax ceftaz1=ceftaz piptazo1=piptazo
  meropenem1=meropenem levomoxi1=levomoxi cipro1=cipro tigecycline1=tigecycline otherbetalactam1=otherbetalactam nonbetalactam1=nonbetalactam;
  *Computing ANYABX where it has missing values;
if record_id in (1983, 2032) then anyabx1=sum(vanco1, linezolid1, daptomycin1, cefotax1, ceftaz1, piptazo1, meropenem1, levomoxi1, cipro1, tigecycline1, otherbetalactam1, nonbetalactam1);
if hospname='West Tennessee Healthcare Milan Hospital' then anyabx1=sum(vanco1, linezolid1, daptomycin1, cefotax1, ceftaz1, piptazo1, meropenem1, levomoxi1, cipro1, tigecycline1, otherbetalactam1, nonbetalactam1);;
*Combining individual antibiotics to make Classes;
  ceph3=sum(cefotax,ceftaz);
 quinolone=sum(levomoxi,cipro);

 *Updating 'Takoma Regional Hospital' and 'Laughlin Memorial Hospital' that changed their names starting in Q2 2019;
if hospname='Takoma Regional Hospital' then hospname='Greeneville Community Hospital';
if hospname='Laughlin Memorial Hospital' then hospname='Greeneville Community Hospital';
*Dropping extreanuous variables;
drop i anyabx vanco linezolid daptomycin cefotax ceftaz piptazo meropenem levomoxi cipro tigecycline otherbetalactam nonbetalactam datacollector
         redcap_survey_identifier   contactphone contactemail specloc___1-specloc___6 icutype___1-icutype___19 othericu  hospid
        stepdown___1-stepdown___3 othericustepdown ward___1-ward___9 otherward hemonc___1-hemonc___7 otherhemonc surveyloc___2
        bonemarrow___1-bonemarrow___3 otherbonemarrow survey_of_antibiotic_v_0 survey_of_antibiotic_v_1;
		output redcap1;
if qtr (surveydate)=&currentq. and year (surveydate)=&currenty. then output redcap2;
run;



/*Import facilities info from NHSN Facility Database + Manipulation*/

PROC IMPORT OUT= WORK.facs
            DATATABLE= 'Facilities_Table_plusbedsize' 
            DBMS=ACCESS REPLACE;     
     DATABASE="H:\NHSN\NHSN Facility Database.accdb"; 
     SCANMEMO=YES;
     USEDATE=NO;
     SCANTIME=YES;
RUN;

data facs (DROP=surveyyear);
set facs (keep=orgid facility_namecat  code2 numbeds surveyyear); 
where	ORGID NE . and  facility_namecat ne " " and surveyyear=2018;*selecting data from the most recent survey year: Update IF change occurs;
if facility_namecat='Takoma Regional Hospital' then facility_namecat='Greeneville Community Hospital West';
if facility_namecat='Laughlin Memorial Hospital' then facility_namecat='Greeneville Community Hospital East';
if facility_namecat='Tennova Healthcare - Volunteer Martin' then do;
                                                                        facility_namecat='West Tennessee Healthcare Volunteer Hospital-Martin';
                                                                                        code2='MT'; 
                                                                                                       end;
if facility_namecat='Milan General Hospital' then facility_namecat='West Tennessee Healthcare Milan Hospital';
if facility_namecat='Camden General Hospital' then facility_namecat='West Tennessee Healthcare Camden Hospital';
facility_namecat=strip(facility_namecat);
rename facility_namecat=hospname;
run;

/* Merging the redcap data with the facility size data by HOSPNAME */

proc sort data=facs; 
by hospname;
proc sort data=redcap2;
by hospname;

data redcap2 (DROP=surveyloc___1); 
	merge redcap2 (in=a) facs;
	by hospname;
	if a;
*Assigning labels to facility sizes ;
	length hospsize $6.;
	if  numbeds < 150 then hospsize="Small";
	else if 150 <= numbeds <= 400 then hospsize="Medium";
	else if numbeds > 400 then hospsize="Large";
	label hospsize="Facility Size";
	* Survey location;
	if surveyloc___1=0 then char="Specific locations";
	else if surveyloc___1=1 then char="Facility-wide"; 
   *Assigning codes to some facilities;
    if hospname='Greeneville Community Hospital East' then code2='AC';
    if hospname='Wayne Medical Center' then code2='CA';
	*Modify location for NorthCrest;
	if record_id=3081 then char="Facility-wide";
run;


*Identify facilities with missing bed size and codes  Then correct their names above;
proc sql;
select record_id, hospname, census 
from redcap2
where orgid is null | numbeds is null | code2 is null;
quit;

*/
******************************************** Current Quarter Analysis **********************************************************************;

/*  Number of surveys table  */
proc freq data=redcap2 noprint;
tables code2/nocum out=freq1(rename=(count=number));
run;

     *creating format for the number of surveys;
proc format;
picture myfmt
  low-high='99 survey(s)';
  run;

      *Creating table for the completion frequency;
  

proc freq data=freq1 noprint;
tables number/nocum out=freq2 ;
run;

  proc sql noprint;
  create table completion as
  select put(number, myfmt.) as Common length=20, count as common2, percent as common3 format=4.1
  from freq2;
  quit;

 

/* Location type table */

proc sql;
create table freq as 
select distinct hospname, char, hospsize
from redcap2;
quit;

proc freq data=freq noprint;
tables char /nocum out=freq3(rename=(char=common count=common2 percent=common3));
run;

/*  Facility Size */

proc freq data=freq noprint;
tables  hospsize/nocum out=freq4(rename=(hospsize=common count=common2 percent=common3));;
run;

/* Concatenating the completion, location and size tables in ONE DATASET */

data collab_table;
set completion freq3 freq4;
*Creating a grouping variable 'tabgroup' for breaking purpose only in PROC REPORT;
if find(common, 'survey') then tabgroup='1';
else if common in ('Large', 'Small', 'Medium') then tabgroup='3';
else tabgroup='2';
run;


******************************************** Summary tables Output **********************************************************************;

/* Percentages of antibiotic use by hospitals for the current quarter */
proc sql; 
title 'Proportions of abx use by hospitals';
create table percentages as
select hospname, avg(census) as av_census 'Average census' format=5.1, sum (anyabx)/sum(census) as new_abx format=percent8.1 'Percentage of anyabx',  sum(vanco)/sum (census) as new_vanco format=percent8.1 'Percentage of Vanco IV', sum (linezolid)/sum (census) as new_linezo format=percent8.1 'Percentage of Linezolid', sum (daptomycin)/sum (census) as new_dapto format=percent8.1 'Percentage of Datomycin',
          sum (ceph3)/sum (census) as new_ceph format=percent8.1 'Percentage of cephalosporines', sum (cefotax)/sum (census) as new_cefo format=percent8.1 'Percentage of Cefotax', sum (ceftaz)/sum (census) as new_ceftaz format=percent8.1 'Percentage of ceftaz', sum (piptazo)/sum (census) as new_piptazo format=percent8.1 'Percentage of Piptazo', sum (meropenem)/sum (census) as new_carbapenem format=percent8.1 'Percentage of carbapenem',sum (quinolone)/sum (census) as new_quinolone format=percent8.1 'Percentage of Quinolone',
		  sum (tigecycline)/sum (census) as new_tige format=percent8.1 'Percentage of tigecycline', sum (otherbetalactam)/sum (census) as new_otherbeta format=percent8.1 'Percentage of Other beta lactam', sum (nonbetalactam)/sum (census) as new_nonbeta format=percent8.1 'Percentage of Non-other Betalactam'
from redcap2
group by 1
order by 1;
quit;

*Collaborative results (MEDIAN MIN MAX) using the dataset generated above;
proc means data=percentages  median min max maxdec=3 noprint; 
title 'Collaborative-wide statistics';
var av_census new_abx new_vanco new_linezo new_dapto new_ceph new_cefo new_ceftaz new_piptazo new_carbapenem new_quinolone new_tige new_otherbeta new_nonbeta ;
output out=stats median= med_new_abx med_new_vanco med_new_linezo med_new_dapto med_new_ceph med_new_cefo med_new_ceftaz med_new_piptazo med_new_carbapenem med_new_quinolone med_new_tige med_new_otherbeta med_new_nonbeta med_av_census
           min= min_new_abx min_new_vanco min_new_linezo min_new_dapto min_new_ceph min_new_cefo min_new_ceftaz min_new_piptazo min_new_carbapenem min_new_quinolone min_new_tige min_new_otherbeta min_new_nonbeta min_av_census
          max= max_new_abx max_new_vanco max_new_linezo max_new_dapto max_new_ceph max_new_cefo max_new_ceftaz max_new_piptazo max_new_carbapenem max_new_quinolone max_new_tige max_new_otherbeta max_new_nonbeta max_av_census;
run;
*Transposing the output statitstics;
proc transpose data=stats (drop=_type_ _freq_) out=transposed  
                                                      name=Molecule prefix=Stat;
run;

*Creating 3 separate datasets: mini maxi and median;
data mini (rename=(stat1=minimum) drop=_label_ molecule ) median (rename=(stat1=median) drop=molecule ) maxi(rename=(stat1=maximum) drop=_label_ molecule);
set transposed;
if index (molecule, 'max') ge 1 then output maxi;
if index ( molecule, 'min') ge 1 then output mini;
if index (molecule, 'med') ge 1 then output median;
run;

*One-to-one merging of the 3 datasets;
data combined;
set median;
set mini;
set maxi;
run;

*Spliting the data in 2: Census & Combined (antibiotics);
data combined census;
set combined;
 if _label_ eq 'Average census' then output census;
 else output combined;
 run;


******************************************** Collaborative wide analysis: Past 6 quarters*******************************************************************************;


 /* Creating collaborative-wide table for each type of abx including all of them (ceftaz, cefotax,tigecycline.. ) this time*/
proc sql; 
create table percentages2 as
select quarter, hospname ,  sum (anyabx)/sum(census) as new_abx format=percent8.1 'Percentage of any abx',  sum(vanco)/sum (census) as new_vanco format=percent8.1 'Percentage of Vanco IV', sum (linezolid)/sum(census) as new_linezo format=percent8.1 'Percentage of Linezo', sum (daptomycin)/sum(census) as new_dapto format=percent8.1 'Percentage of Dapto',
          sum (ceph3)/sum (census) as new_ceph format=percent8.1 'Percentage of cephalosporines', sum (cefotax)/sum(census) as new_cefo format=percent8.1 'Percentage of cefotax', sum (ceftaz)/sum(census) as new_ceftaz format=percent8.1 'Percentage of ceftaz', sum (piptazo)/sum (census) as new_piptazo format=percent8.1 'Percentage of Piptazo', sum (meropenem)/sum (census) as new_carbapenem format=percent8.1 'Percentage of carbapenem',
          sum (quinolone)/sum (census) as new_quinolone format=percent8.1 'Percentage of Quinolone', sum (tigecycline)/sum(census) as new_tige format=percent8.1 'Percentage of Tigecycline', sum (otherbetalactam)/sum(census) as new_otherbeta format=percent8.1 'Percentage of Other betalactam', sum (nonbetalactam)/sum(census) as new_nonbeta format=percent8.1 'Percentage of Non betalactam', sum(census)/count(census) as new_census format=4. "Average census"
from redcap1
group by 1, 2
order by 1 desc, 2;
quit;

*Transposing data to make "quarter" values as variables and "types of abx" as rows;
proc sort data=percentages2 out=sorted;
by hospname;run;
proc transpose data=sorted out=transposed2 name=Abx ;
id quarter;
by hospname;
run;

/* Format for the molecules & census Colums */
proc format;
value $ molecul
   'Percentage of any abx'='Any antibiotic'
   'Percentage of Vanco IV'='Vancomycin IV'
   'Percentage of Linezo'='Linezolid'
   'Percentage of Dapto'='Daptomycin'
   'Percentage of cephalosporines'='Cephalosporins combined'
   'Percentage of cefotax'='Non-antipseudomonal 3G cephalosporins'
   'Percentage of ceftaz'='Antipseudomonal cephalosporins'
   'Percentage of Piptazo'='Piperacillin/tazobactam'
   'Percentage of carbapenem'='Carbapenems'
   'Percentage of Quinolone'='Quinolones'
   'Percentage of Tigecycline'='Tigecycline'
   'Percentage of Other betalactam'='Any other beta-lactam'
   'Percentage of Non betalactam'='Any other non-beta-lactam';
   run;
   proc format;
   value $census
   'Average census'='Census at 9:00 AM(or other specified time) on survey date';
   run;



********************************************BAR-LINE graphs Data Prep ************************************************************************************************;

   /* Percentages of abx use by quarter & facility: facility specific dataset */
proc sql; 
create table percentages3 as
select quarter, hospname , sum (anyabx)/sum(census) as new_abx format=percent8.1 'Percentage of anyabx',  sum(vanco)/sum (census) as new_vanco format=percent8.1 'Percentage of Vanco IV', 
          sum (ceph3)/sum (census) as new_ceph format=percent8.1 'Percentage of cephalosporines', sum (piptazo)/sum (census) as new_piptazo format=percent8.1 'Percentage of Piptazo', sum (meropenem)/sum (census) as new_carbapenem format=percent8.1 'Percentage of carbapenem',sum (quinolone)/sum (census) as new_quinolone format=percent8.1 'Percentage of Quinolone'
from redcap1
group by 1, 2
order by 1, 2;
quit;

 /*Creating collaborative-wide results using the previous dataset created */
proc sql;
create table median as
select quarter, median(new_abx) as median_anyabx format=percent8.1, median (new_vanco) as median_vanco format=percent8.1 , median(new_ceph) as median_cef format=percent8.1,
median(new_piptazo) as median_piptazo format=percent8.1, median(new_carbapenem) as median_carba format=percent8.1, median( new_quinolone) as median_quino format=percent8.1
from percentages3
group by 1
order by 1;
quit;

  /*Right join of collaborative-wide with facility specific */
proc sql;
create table merged as
select * 
from percentages3 p right join median  m
on p.quarter=m.quarter;
quit;


/* Adding missing values of the collab-wide results row for Facilities that did not report in any of the past 6 quarters
     Purpose: To have the quarter for possible missing info to be represented on the graph with only collab results */

proc sql noprint;
select hospname into: hospLT6 separated by  "/" 
from (select hospname, count (quarter) as count
            from merged where hospname in (select distinct hospname from redcap2 )
           group by 1)
where count<6;
select count( distinct hospname) into: counthosp
from (select hospname, count (quarter) as count
            from merged where hospname in (select distinct hospname from redcap2 )
           group by 1)
where count<6;
quit;
%put %bquote(&hospLT6) &counthosp;

* Template from Merged;
proc sql;
create table temp as
select distinct quarter, median_anyabx , median_vanco ,median_cef ,median_piptazo ,median_carba , median_quino
from merged;
quit;

data temp;
set temp;
new_abx=.; 
new_vanco=.; 
new_ceph=.;
new_piptazo=.;
new_carbapenem=.;
new_quinolone=.;
run;

/* Macro to process Missing values*/


%macro missing;
%do i=1 %to &counthosp;
  *Adding the missing hospname variable ;

data temp&i;
set temp;
hospname="%scan(%bquote(&hospLT6), &i, /)";
run;
 
   *Updating the new data with their existing values in the merged/using the update statement;
data updated&i;
update temp&i merged (where=(hospname="%scan(%bquote(&hospLT6), &i, /)"));
by quarter hospname;
run;
    *Concatenate the upadated facility data with the big set 'Merged';
data merged ;
set merged updated&i;
run; 
%end;
%mend ;

*Execute the macro program;
%missing;

*Delete duplicated rows;

proc sort data=merged nodup;
by quarter hospname;
run;



************************************ MACRO DEFINITION FOR PDF OUTPUTS **************************************************************************************;


 ods noproctitle escapechar='^';
options nodate nonumber papersize=standard orientation=portrait    topmargin=.1in bottommargin=0in leftmargin=.2in rightmargin=.1in;	

         /* Plots Macro To be Nested in The Final macro Program */
%macro plot (abxbar,abxline, label);

proc sgplot data=merged;
where hospname="%bquote(&hospname)";
vbar quarter/response=new_&abxbar datalabel legendlabel='Facility' fill fillattrs=(color=vligb);
vline quarter/ response=median_&abxline markers markerattrs=(color=black symbol=circlefilled size=0.1in) lineattrs=(color=blue thickness=0.05in) nostatlabel legendlabel='Collab Wide';
yaxis grid label="Antibiotic Use Proportion";
xaxis label='Quarter';
title " ^S={font_size=0.5}&label Use percent - Collaborative Wide vs. Facility Specific";
run;

%mend;

           /* Proc Report Macro To be Nested to the Final Macro Program */
*Macro for variable names to be supplied in the proc report;
proc contents data=transposed2 noprint out=vars (keep=name);
run;
proc sql noprint;
select name into: var separated by '/'
from vars
where name like '%_20%'
order by 1 desc;
quit;
%put &var;

*Macro for quarters and years;
proc sql noprint;
select distinct quarter   /* Qt=last 6 quarters */
           into: qt separated by '/'
from redcap1
order by 1 desc;
select distinct substr(quarter,1,4)  /* Yr=Years covered by the last 6 quarters */
           into: yr separated by '/'
from redcap1
order by 1 desc;
select name into: var1 separated by '  '  /* Var1=Variables for the most recent year */
from vars
where input(substr(name,2,4), best12.) = %scan(&yr,1,/)
order by 1 desc;
select name into: var2 separated by '  '  /* Var2=Variables for the second most recent year */
from vars
where input(substr(name,2,4), best12.) = %scan(&yr,2,/)
order by 1 desc;
select name into: var3 separated by '  '  /* Var3=Variables for the oldest year */
from vars
where input(substr(name,2,4), best12.) = %scan(&yr,3,/);
quit;
%put &var2;


%macro report (data, title, format,type, format2);

%if &currentq=1 %then %do;

proc report data=&data  split='#' nowindows spanrows
  style(report)=[fontfamily="Calibri" borderwidth=2 bordercolor=black textalign=c]
  style(header)=[color=black  fontfamily="Calibri " fontsize=2 fontweight=bold  backgroundcolor=lightgrey foreground=Black borderwidth=3   textalign=c]
  style(column)=[color=black  fontfamily="Calibri" fontsize=2 fontweight=medium borderwidth=3 textalign=c]  
  style(lines)=[color=black   fontfamily="Arial" fontsize=1 fontweight=bold backgroundcolor=lightgrey textalign=c  ]
  style(summary)=[color=black fontfamily="Arial" fontsize=1 fontweight=bold backgroundcolor=lightgrey textalign=c];

title "Summary Table of &title &quarter.";
                     *UPDATE below to reflect the last 6 quarters: on the COLUMN line;
  column  _LABEL_ ("Collaborative Wide Results of Quarter &currentq.- &currenty."  median minimum maximum) 
                                    ("Facility &code" ("%scan(&yr,1,/)" &var1.  ) ("%scan(&yr,2,/)" &var2. ) ("%scan(&yr,3,/)" &var3.  )); 	*UPDATE years macro;				
	
  DEFINE _LABEL_/DISPLAY  FORMAT=$&format.. STYLE(COLUMN) = [JUST=left width=2.0in] "&type";
  DEFINE median/DISPLAY FORMAT=&format2  'Median';
  DEFINE minimum/DISPLAY FORMAT=&format2  'Minimum';
  DEFINE maximum  /DISPLAY FORMAT=&format2 'Maximum' ;
  DEFINE %scan(&var,1,/) /DISPLAY FORMAT=&format2   "Quarter %substr(%scan(&qt,1,/),7,1)"; 
  DEFINE %scan(&var,2,/) /DISPLAY FORMAT=&format2 "Quarter %substr(%scan(&qt,2,/),7,1)" ;
  DEFINE %scan(&var,3,/) /DISPLAY FORMAT=&format2 "Quarter %substr(%scan(&qt,3,/),7,1)";  
  DEFINE %scan(&var,4,/) /DISPLAY FORMAT=&format2 "Quarter %substr(%scan(&qt,4,/),7,1)";
  DEFINE %scan(&var,5,/) /DISPLAY FORMAT=&format2  "Quarter %substr(%scan(&qt,5,/),7,1)";
  DEFINE %scan(&var,6,/) /DISPLAY FORMAT=&format2  "Quarter %substr(%scan(&qt,6,/),7,1)";
  RUN;


                       %end;


                  %else %do;

  proc report data=&data  split='#' nowindows spanrows
  style(report)=[fontfamily="Calibri" borderwidth=2 bordercolor=black textalign=c]
  style(header)=[color=black  fontfamily="Calibri " fontsize=2 fontweight=bold  backgroundcolor=lightgrey foreground=Black borderwidth=3   textalign=c]
  style(column)=[color=black  fontfamily="Calibri" fontsize=2 fontweight=medium borderwidth=3 textalign=c]  
  style(lines)=[color=black   fontfamily="Arial" fontsize=1 fontweight=bold backgroundcolor=lightgrey textalign=c  ]
  style(summary)=[color=black fontfamily="Arial" fontsize=1 fontweight=bold backgroundcolor=lightgrey textalign=c];

   title "Summary Table of &title &quarter.";
                     *UPDATE below to reflect the last 6 quarters: on the COLUMN line;
  column  _LABEL_ ("Collaborative Wide Results of Quarter &currentq.- &currenty."  median minimum maximum) 
                                    ("Facility &code" ("%scan(&yr,1,/)" &var1.  ) ("%scan(&yr,2,/)" &var2. ) ); 	*UPDATE years macro;				
	
  DEFINE _LABEL_/DISPLAY  FORMAT=$&format.. STYLE(COLUMN) = [JUST=left width=2.0in] "&type";
  DEFINE median/DISPLAY FORMAT=&format2  'Median';
  DEFINE minimum/DISPLAY FORMAT=&format2  'Minimum';
  DEFINE maximum  /DISPLAY FORMAT=&format2 'Maximum' ;
  DEFINE %scan(&var,1,/) /DISPLAY FORMAT=&format2   "Quarter %substr(%scan(&qt,1,/),7,1)"; 
  DEFINE %scan(&var,2,/) /DISPLAY FORMAT=&format2 "Quarter %substr(%scan(&qt,2,/),7,1)" ;
  DEFINE %scan(&var,3,/) /DISPLAY FORMAT=&format2 "Quarter %substr(%scan(&qt,3,/),7,1)";  
  DEFINE %scan(&var,4,/) /DISPLAY FORMAT=&format2 "Quarter %substr(%scan(&qt,4,/),7,1)";
  DEFINE %scan(&var,5,/) /DISPLAY FORMAT=&format2  "Quarter %substr(%scan(&qt,5,/),7,1)";
  DEFINE %scan(&var,6,/) /DISPLAY FORMAT=&format2  "Quarter %substr(%scan(&qt,6,/),7,1)";
  RUN;

                       %end;

  %mend ;


 /* Macro Program to output PDF formattted Final Report for each Facility */

 %macro pdf (hospname, code);
         title;
                  *Defining the output path for the pdf document;
ods pdf file="&path.\%bquote(&hospname).pdf"  NOTOC startpage=never dpi=300;

                  *Defining the Cover Page;

ods layout start width=8in height=10.9in; *Page 1;
ods region ;
ods region x=0.5in y=0.1in  ;

ods pdf text='^S={preimage="H:\NHSN\Antimicrobial Stewardship\TDH Logo.png" just=c}';
ods pdf text=" ";
ods pdf text=" ";
ods pdf text="^S={font_weight=bold textalign=c font_size=20pt  font_face='Arial Black' color=black textdecoration=underline}Antimicrobial Use Survey Report for &quarter";
ods pdf text=" ";
ods region x=0.5in y=5in  ;
ods pdf text="^S={font_weight=bold textalign=c font_size=16pt  font_face='Arial Black' color=black} %bquote(&hospname)";
ods layout end;

              *starting the new page and layout;
ods startpage=now;*Page 2;         
ods layout start;

              *defining region for the first table;
ods region x=0.5in y=1in;

%* Creating macro variables ;

           *Macro variable for location type;
proc freq data=freq noprint;
tables char/nocum out=char;
where hospname="%bquote(&hospname)";

data _null_;
set char;
call symputx ('location', char);
run;

              *Macro variable for fac size;

proc freq data=freq noprint;
tables hospsize/nocum out=size;
where hospname="%bquote(&hospname)";

data _null_;
set size;
call symputx ('size', hospsize);
run;

            *macro variables to match code & number of suveys;


data _null_;
set freq1;
if code2="&code" then call symputx ('survey', number);


%*Proc Report for the first table ;

proc report data=collab_table nowindows spanrows headskip
  style(report)=[fontfamily="Calibri" borderwidth=2 bordercolor=black textalign=c]
  style(header)=[color=black  fontfamily="Calibri " fontsize=2 fontweight=bold  backgroundcolor=lightgrey foreground=Black borderwidth=3   textalign=c]
  style(column)=[color=black  fontfamily="Calibri" fontsize=2 fontweight=medium borderwidth=3 textalign=c]  
   style(lines)=[color=black   fontfamily="Arial" fontsize=2  backgroundcolor=lightgrey  ];

title 'Collaborative Wide and Facility Specific Results';

column tabgroup ('Collaborative Wide Results' common common2 common3) ;

define tabgroup/group noprint ;
break after tabgroup / summarize dol style=[fontweight=bold  backgroundcolor=lightgrey  ];
define common/order 'Characteristic'  STYLE(COLUMN) = [JUST=center width=2in];
define common2/ 'N' center STYLE(COLUMN) = [JUST=center width=2in];
define common3/  '%' STYLE(COLUMN) = [JUST=center width=2in];
compute after;
line @40  "^S={ font_weight=bold font_size=3 textdecoration=underline } Facility Specific Results:";
line "  ";
line @43 "^S={font_size=2 }Your facility size: ^S={font_weight=bold font_size=2 } &size";
line @43 "^S={font_size=2 }Your location type: ^S={font_weight=bold font_size=2 } &location";
line @43 "^S={font_size=2 }Number of surveys submitted by your facility: ^S={font_weight=bold font_size=2 } &survey";
endcomp;
run;

title;
               *Ending the 2nd page layout;
ods layout end;

              *Starting the 3rd page;
ods startpage=now;
ods layout start;
ods region   x=0.1 ;

              *Subsetting the specific facility data, whithout Census row. Reminder: data Transposed2 was created previously  ;
   data fac_data;
   set transposed2;
   where hospname="%bquote(&hospname)" & abx ne "new_census";

              *One-to-One merging of facility-wide data for current quarter + facility-specific previous quarter results;
   data final;
   set combined;*Reminder: the data combined was created previously;
   set fac_data;
   run;

   %* Proc Report for the Antibiotics table of the second page ;

     %report (final, Antibiotics, molecul, Molecule, percent7.1);
  
   %*Census Table ;

  ods region y=6.7 in;

data census_data(drop=abx hospname);*Subsetting the specific facility census data;
set transposed2;
where hospname="%bquote(&hospname)" & abx eq "new_census";

data final_census;*One-to-One merging of facility-wide census data for current quarter + facility-specific previous quarter results;
set census;
set census_data;
run;
    %*Proc Report for the Census table of the second page ;

%report (final_census, Census, census, Census, 5.1);

   title;

ods layout end;*Ending the 3rd page;


ods startpage=now;*Starting the 4th page;
ods layout start;


%* Defining  bar-line chart Regions  for every molecule. Order= Vanco-Ceftri-Carbapenem-Quinolone-Any Abx-Piptazo ;
ods region width=3.5in height=3.7in x=0.2in y=0.2in;

%plot (vanco, vanco, Vancomycin);

ods region width=3.5in height=3.7in x=3.9in  y=0.2in;

%plot (ceph, cef, Cephalosporins);

ods region width=3.5in height=3.7in x=0.2in y=3.4in;

%plot (carbapenem, carba, Carbapenems);

ods region width=3.5in height=3.7in x=3.9in y=3.4in;

%plot (quinolone, quino, Quinolones);

ods region width=3.5in height=3.7in x=0.2in y=6.7in;

%plot (abx, anyabx, Any Antibiotic);

ods region width=3.5in height=3.7in x=3.9in y=6.7in;

%plot (piptazo, piptazo, Pip/Tazo);

ods layout end;

ods pdf close;

%mend ;


/* Invoking Macro for each participating facility */


%macro Output;

proc sql noprint;
select distinct hospname, code2, count (distinct hospname) into : hospi separated by ' / ' ,
                                                                                                                 :cod separated by ' / ' , 
                                                                                                                  :n
from redcap2;
quit;

%do i=1 %to &n;

%pdf (%scan (%bquote(&hospi),&i,/), %scan(&cod,&i, /));

%end;

%mend;

   *The below macro call will generate the output;
%output;
Clear Log Window ;
 DM "log; clear; ";*/
