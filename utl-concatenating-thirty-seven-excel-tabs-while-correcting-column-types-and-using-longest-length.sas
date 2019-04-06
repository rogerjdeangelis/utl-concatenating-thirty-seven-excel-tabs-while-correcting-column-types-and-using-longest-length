Concatenating thirty seven excel tabs while correcting column types and using longest lengths                                           
                                                                                                                                        
github                                                                                                                                  
http://tinyurl.com/y4ferymu                                                                                                             
https://github.com/rogerjdeangelis/utl-concatenating-thirty-seven-excel-tabs-while-correcting-column-types-and-using-longest-length     
                                                                                                                                        
I was given a workbook with 37 tabs and I needed to concatenate all 37.                                                                 
There is a common set of columns in all 37 tabs, however some columns were                                                              
numeric in one tab and character in another.                                                                                            
both numeric and charater.                                                                                                              
                                                                                                                                        
I dont use 'proc import' for obvious reasons.                                                                                           
                                                                                                                                        
First you need to set guessing rows to the max 32756 by changing a registry setting.                                                    
                                                                                                                                        
https://www.lexjansen.com/pharmasug-cn/2014/PT/PharmaSUG-China-2014-PT09.pdf                                                            
                                                                                                                                        
4. How to add row numbers which determine one variable’s type                                                                           
TypeGuessRows                                                                                                                           
In SAS excel engine, we need to change the windows registry settings, the locations are not same in each windows                        
or MS office version.                                                                                                                   
For windows7 x64 with MS office 2013 64-bit, it is in                                                                                   
HKEY_LOCAL_MACHINE / Software / Microsoft /Office/15.0 /AccessConnectivityEngine /Engines/ Excel                                        
Change the value of TypeGuessRows to 0, which by default, is 8 in hexadecimal.                                                          
Then all the rows will be scanned and checked.                                                                                          
But we should know that changes also affect other software that uses the Microsoft Jet provider to access Excel file                    
data, including accessing Excel data in a Microsoft Access database.                                                                    
                                                                                                                                        
                                                                                                                                        
Here is the solution                                                                                                                    
                                                                                                                                        
                                                                                                                                        
* copy all 37 to work datasets - note the option on the libname. Perhaps you don't need to edit the registry                            
  scantext=yes scans the entire column??;                                                                                               
libname dna "d:/dna";                                                                                                                   
libname xel "d:/mbs/xls/data_definition_table.xlsx" dbMax_text=32767 mixed=yes scantext=yes;                                            
                                                                                                                                        
proc copy in=xel out=work;                                                                                                              
run;quit;                                                                                                                               
                                                                                                                                        
libname xel clear;                                                                                                                      
                                                                                                                                        
* get the names of the imported tables;                                                                                                 
                                                                                                                                        
proc sql;                                                                                                                               
  select                                                                                                                                
    memname                                                                                                                             
  from                                                                                                                                  
      sashelp.vtable                                                                                                                    
  where                                                                                                                                 
    libname ="WORK                                                                                                                      
;quit;                                                                                                                                  
                                                                                                                                        
* manually paste the names into this array macro leaving off the first work daatset;                                                    
%array(dats,values=                                                                                                                     
JCQJ JDLS CZRZ CZSJ CZST DENT ENUM EVNT EVZS HERZ HEST                                                                                  
HLPR HZME HZUS HRND IJDL INCZ INTV IRQS IRQS KNZW MZBL MPQF NJQS ZSZP PJYM                                                              
PDRX PLJN PLRZ PMEV PMQF PMRZ  PRZV PVQS RZST USQW XCEV);                                                                               
                                                                                                                                        
* delete the concatenation table if it exists;                                                                                          
                                                                                                                                        
proc datasets lib=mbs noprint;                                                                                                          
 delete allcdebok                                                                                                                       
run;quit;                                                                                                                               
                                                                                                                                        
%array(dats,value=&mems);                                                                                                               
                                                                                                                                        
* here is a manual example;                                                                                                             
                                                                                                                                        
* Two important properties of SQL                                                                                                       
   1. It can convert numeric variables to character with the same name.                                                                 
   2. Union will use the longets length found in any of the work table.                                                                 
;                                                                                                                                       
                                                                                                                                        
proc sql;                                                                                                                               
  create                                                                                                                                
     table work.allcdebok as                                                                                                            
  select                                                                                                                                
      cats(ANALYTIC_CODE_FRAME     )   as ANALYTIC_CODE_FRAME      length=4096                                                          
     ,cats(ANALYTIC_DATASET        )   as ANALYTIC_DATASET         length=4096                                                          
     ,cats(ANALYTIC_FORMAT_NAME    )   as ANALYTIC_FORMAT_NAME     length=4096                                                          
     ,cats(ANALYTIC_LABEL          )   as ANALYTIC_LABEL           length=4096                                                          
     ,cats(ANALYTIC_VARIABLE_LENGTH)   as ANALYTIC_VARIABLE_LENGTH length=4096                                                          
     ,cats(ANALYTIC_VARIABLE_NAME  )   as ANALYTIC_VARIABLE_NAME   length=4096                                                          
     ,cats(ANALYTIC_VARIABLE_TYPE  )   as ANALYTIC_VARIABLE_TYPE   length=4096                                                          
  from                                                                                                                                  
     jxxcx                                                                                                                              
  %do_over(dats,phrase=%str(                                                                                                            
  union                                                                                                                                 
     corr                                                                                                                               
  select                                                                                                                                
      cats(ANALYTIC_CODE_FRAME     )   as ANALYTIC_CODE_FRAME      length=4096                                                          
     ,cats(ANALYTIC_DATASET        )   as ANALYTIC_DATASET         length=4096                                                          
     ,cats(ANALYTIC_FORMAT_NAME    )   as ANALYTIC_FORMAT_NAME     length=4096                                                          
     ,cats(ANALYTIC_LABEL          )   as ANALYTIC_LABEL           length=4096                                                          
     ,cats(ANALYTIC_VARIABLE_LENGTH)   as ANALYTIC_VARIABLE_LENGTH length=4096                                                          
     ,cats(ANALYTIC_VARIABLE_NAME  )   as ANALYTIC_VARIABLE_NAME   length=4096                                                          
     ,cats(ANALYTIC_VARIABLE_TYPE  )   as ANALYTIC_VARIABLE_TYPE   length=4096                                                          
  from                                                                                                                                  
    work.?                                                                                                                              
  ));                                                                                                                                   
quit;                                                                                                                                   
                                                                                                                                        
                                                                                                                                        
work.allcdebook is the concatenation of all 37 tabbs.                                                                                   
                                                                                                                                        
Note column lengths can be longer that 1024.                                                                                            
                                                                                                                                        
                                                                                                                                        
                                                                                                                                        
