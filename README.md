# AFM-plots_Jython
Five acceptance criteria
A curve is accepted only when all five conditions hold:
1. Custom linear fit RMSD ≤ 25 pN.
2. JPK elasticity fit RMS ≤ 25 pN.
3. Full linear fitted X interval width ≥ 5 nm.
4. Elasticity fitted interval on the indentation side of the processed
   contact at X = 0 nm is ≥ 10 nm. The interval is clipped to the
   exported elasticity fit limits.
5. Length of the overlapping X interval between the linear and elasticity
   fits ≤ 5 nm.