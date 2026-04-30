Usage of Excel sheet for scheduling and scoring are as follows,

1) Register everyone's attendance in the sheet "attendance" as '1'. If any
player inform late, enter their ETA in minutes, instead of '1'. If any player is
absent, mark their attendance as '0'.

2) Once enough teams have arrived, run python script. Python script fill sheets
"quali_schedule" and "quali_score".

3) Share "quali_schedule" sheet to all teams so that they know their turns.

4) Every match results needs to be entered in the sheet "quali_score".

5) Once all qualification rounds are complete, copy group quali results at the
end of 'quali_score' sheet into 'quali_results' sheet by values only and aligned
with column 'A' entries of 'quali_results'.

6) In 'quali_results' sheet, every group's results need to be sorted according to
points and score.

7) Once sorted in step 6), correct player list will be displayed in sheet
'knockout_score'.

8) Complete 'knockout_score' sheet manually.


Python script schedule the match as follows,
1) It will attempt to fill first two rounds with attendance registered teams.

2) No match is scheduled successively for a team.

3) All the Boys & Girls matches are on Court 4.

4) Court 5 will be used only only once for a team.

5) Matches are chosen randomly.

6) There won't be any schedule for absent teams.

Useful commands
libreoffice --invisible --convert-to html ./bracket.xlsx --outdir /tmp/aa
xlsx2csv -s 1 ./bracket.xlsx attendance.csv

=if(rand()>=0.5,21,int(rand()*20))
=if(C4=21,int(rand()*20),21)
