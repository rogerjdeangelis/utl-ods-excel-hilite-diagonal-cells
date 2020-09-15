# utl-ods-excel-hilite-diagonal-cells
Ods excel hilite diagonal cells
    Ods excel hilite diagonal cells                                                                         
                                                                                                            
    github                                                                                                  
    https://github.com/rogerjdeangelis/utl-ods-excel-hilite-diagonal-cells                                  
                                                                                                            
    Excel output for this repo                                                                              
    https://tinyurl.com/y5qju2k2                                                                            
    https://github.com/rogerjdeangelis/utl-ods-excel-hilite-diagonal-cells/blob/master/want.xlsx            
                                                                                                            
    SAS Forum                                                                                               
    https://tinyurl.com/y67mnmea                                                                            
    https://communities.sas.com/t5/SAS-Procedures/proc-report-tabulate-and-matrix-diagonal/m-p/683576       
                                                                                                            
    FYI: You can preprocess and apply FOREGROUND colors hiliteing (just colors the text)                    
                                                                                                            
         if row=col then do; cell=cats("^S={foreground=black}",col);                                        
                                                                                                            
         But unfortunately you CANNOT do BACKGROUND hiliteing                                               
                                                                                                            
         if row=col then do; cell=cats("^S={background=black}",col);                                        
         ( this would be much simpler does not work in clssic SAS 9.4M6 64bit Win 10 64bit )                
         ( maybe it will work with one of the Frankenstein SAS implementations. I don't use EG)             
                                                                                                            
    Related github repos                                                                                    
                                                                                                            
    https://tinyurl.com/y6bm4ytv                                                                            
    https://github.com/rogerjdeangelis/utl_excel_experimenting-with-the-new-ods-excel-destination           
                                                                                                            
    https://tinyurl.com/y29dmhds                                                                            
                                                                                                            
                                                                                                            
    /*                   _                                                                                  
    (_)_ __  _ __  _   _| |_                                                                                
    | | `_ \| `_ \| | | | __|                                                                               
    | | | | | |_) | |_| | |_                                                                                
    |_|_| |_| .__/ \__,_|\__|                                                                               
            |_|                                                                                             
    */                                                                                                      
                                                                                                            
    options ps=16 ls=64;                                                                                    
    data have;                                                                                              
     array cols[3] $32 c1-c3 ('1','2','3');                                                                 
      do row=1 to 3;                                                                                        
       do col=1 to 3;                                                                                       
         cols[col]=col;                                                                                     
          if row=col then cols[col]=col!!" ";                                                               
       end;                                                                                                 
       drop row col;                                                                                        
       output;                                                                                              
      end;                                                                                                  
      stop;                                                                                                 
    run;quit;                                                                                               
                                                                                                            
    WORK.HAVE                                                                                               
                                                                                                            
      C1        C2        C3                                                                                
                                                                                                            
       1         1         1                                                                                
       2         2         2                                                                                
       3         3         3                                                                                
                                                                                                            
    /*           _               _                                                                          
      ___  _   _| |_ _ __  _   _| |_                                                                        
     / _ \| | | | __| `_ \| | | | __|                                                                       
    | (_) | |_| | |_| |_) | |_| | |_                                                                        
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                       
                    |_|                                                                                     
    */                                                                                                      
                                                                                                            
    --------------------------------------                                                                  
    |   |   C1     |   C2     |   C3     |                                                                  
    |------------------------------------|                                                                  
    |  1|    1     |    1     |    1     |                                                                  
    |   | YELLOW   |          |          | * col 1 row 1 background=yellow                                  
    |---+----------+----------+----------|                                                                  
    |  2|    2     |    2     |    2     |                                                                  
    |   |          | YELLOW   |          |                                                                  
    |---+----------+----------+----------|                                                                  
    |  3|    3     |    3     |    3     |                                                                  
    |   |          |          |  YELLOW  |                                                                  
    --------------------------------------                                                                  
                                                                                                            
    [WANT]                                                                                                  
                                                                                                            
    /*                                                                                                      
     _ __  _ __ ___   ___ ___  ___ ___                                                                      
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|                                                                     
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                     
    | .__/|_|  \___/ \___\___||___/___/                                                                     
    |_|                                                                                                     
    */                                                                                                      
                                                                                                            
    ods escapechar="^";                                                                                     
    ods excel file="d:/xls/want.xlsx";                                                                      
    ods excel options(sheet_name="WANT");                                                                   
    proc report data=have style(column)={PROTECTSPECIALCHARS=off};                                          
    cols c1 c2 c3;                                                                                          
    define c1 / display right;                                                                              
    define c2 / display right ;                                                                             
    define c3 / display right ;                                                                             
    compute c1 ;                                                                                            
       if index(c1," ") then call define(_col_,"style","style={just=right background=yellow}");             
    endcomp;                                                                                                
    compute c2 ;                                                                                            
       if index(c2," ")  then call define(_col_,"style","style={just=right background=yellow}");            
    endcomp;                                                                                                
    compute c3 ;                                                                                            
       if index(c3," ")  then call define(_col_,"style","style={just=right background=yellow}");            
    endcomp;                                                                                                
    run;quit;                                                                                               
    ods excel close;                                                                                        
                                                                                                            
                                                                                                            
                                                                                                            
