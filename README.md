# p2tmaker

creates a p2t file out of the information from the 
 + PRZM result file (*.zts)
    + content      : RUNF, PRCP, RFLX1, RFLX2, ESLS, EFLX1, EFLX2, TPAP, IRRG
    + file name    : Scenario, Crop
 + MASTER.FPJ file : chemical names
 + przm.pzm file   : SWASH numbers

Calculation of GW storage and discharge 

+ GW_discharge calc. with Stella from Nick Jarvis

   + GW_storage = GW_storage + ((INFL - GW_discharge) * Timestep)
   + GW_discharge = (1 / MRT) * GW_storage
   
  
+ GW_discharge  calc. with exponential discharge formula
  Q2 = Q1 exp { −A (T2 − T1) } + R [ 1 − exp { −A (T2 − T1) } ]
  https://en.wikipedia.org/wiki/Runoff_model_(reservoir)

   + GW_discharge =
      GW_discharge * Math.Exp(-(1 / MRT) * Timestep) +
      INFL * (1 - Math.Exp(-(1 / MRT) * Timestep))  
 
+ with  MRT = Mean residence time in days, std. = 20d

-------------------------------------------------

## Usage

convert a **single** zts file
 + zts:='full zts file path'
 
**OR**

convert **all** zts files in a project directory
 + path:='full path to directory with zts files'

get ZTS files recursively, **std. = false**
 + recursive:=true/false

if 'p2tmaker.exe' is in the project directory
it can be used without these two cmd line args

To skip warm up years, **default = 0**
 + warmup:=0

To set the mean residence time in days, **default = 20days**
 + mrt:=20

To set the max. precipitation per hour for
calculation of event duration in mm, **default = 2mm**
 + maxPREC:=2

GW_discharge calc. with exponential discharge formula, **std. = false**
 + exp:=true/false

