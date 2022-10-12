
*#################################################################

Objectif : 		Réalisation des analyses statistiques "standards"
 				par les membres du service ESSAI

Auteur : 		Cédric JUNG

Date :			30/12/2019

Version : 		0.3.3

#################################################################
;
****************************************************************
1.lang fr vs ANG
2.TEST COLONNE
3.TITRE ET METHODE FR VS ANG
4.TABLEAU PAR PARAMETER
5.GRAPHIQUE PAR TIME<1000 ET TIME>1000
6.TABLEAU VARIATION % [MEAN(Di)- MEAN(D0)] /MEAN(D0)
****************************************************************;

ODS escapechar='^';

ods graphics on / 
      width=3in
      imagefmt=jpeg
      imagemap=on
      imagename="MyBoxplot"
      border=off;

*DECLARATION DES MACRO-VARIABLES;

%let table=_22E0794;
%let nombre_decimales=2;
%let plan_experimental=Intra_individuel;	
*%let plan_experimental= Groupes_paralleles;
%let lang= ANG;  /* ang/fr */

*--------------------------------------
Liste des macro-variables à renseigner
pour les sous programmes
;

%let 	delta		= abs 	;
%let 	autresVariables = 	; 
%let	libFormat 	= 		;
%let	tableVar 	=  		;
%let 	tableSortie = all	;
%let 	tableEntree= temp._&table ; 
%let 	num_etude = &table;


/***Création d'un nouveau style "template"***/

proc template;
	define style sty.Eurofins1;
	parent=styles.rtf;
/**************************************
	test pour modifier les marges
**************************************/
	replace Body from Document /
	leftmargin = 3 cm
	rightmargin = 3 cm
	topmargin = 1.2 cm
	bottommargin = 1.2 cm;

	style systemtitle /
	font_face = "arial, helvetica, sans-sherif"
	font_size = 10pt
	font_weight = bold
	font_style= italic
	background=white;

	style header /
	background=#D9D9EA
	color=#5E64A4      /*for anna  mgu*/
	font_weight = bold
	font_size =10pt
	font_face = "calibri"
	; 

	style rowheader /
	font_weight = bold
	font_size =10pt
	font_face = "calibri";

	style footer /
	background=#E2FDFE
	font_weight = bold
	font_size =10pt
	font_face = "calibri"
	;

	style table /
	cellpadding=3
	cellspacing=0
	bordercolor=black
	foreground=black
	just=center
	rules = groups
	frame = hsides; 

	style column  /
	just=c;

	style proctitle /
	foreground=#1C3583
	font_weight = bold
	font_size =10pt
	font_face = "calibri";

	style data /
	font_size =10pt
	font_face = "calibri";
end;
run;	





*--------------------------------------------
IMPORTATION DES DONNEES
Création de la table SAS à partir du fichier 
Excel
;

libname temp "d:/Techniciens/Macro_SAS/";

proc import out=temp._&table
DATAFILE="D:\Techniciens\MACRO_SAS\&table..xlsx"
DBMS=XLSx REPLACE;
RUN;


*Suppression des lignes vides;
ods output position=_POSITION;
proc contents data=&tableEntree order=varnum;
run;

data _POSITION;
	set _POSITION;
	if Num = 1 then do; 
		call symput('colonne_produit',variable); end;
	else if Num = 2 then do;
		call symput('colonne_subject',variable); end;

%put 
colonne_produit: &colonne_produit
colonne_subject: &colonne_subject;
run;

data temp._&table;
	set temp._&table;
	if not missing(&colonne_produit);
	if not missing(&colonne_subject);
run;





*------------------------------------------------

EXTRACTION et STCOCKAGE dans des macro-variables
des information suivantes
	- variable produit
	- variable sujet
	- liste des variables à analyser
;

ods trace on;
ods output position=_POSITION;
proc contents data=&tableEntree order=varnum;
run;

proc sql;
	select variable, type
	into : product, : type_variable_produit
	from _POSITION
	where num = 1
;
	select variable
	into : subject
	from _POSITION
	where num = 2
;
	select variable
	into : varClassement separated by ' '
	from _POSITION
	where num < 3
;
	select variable
	into : varTransposees separated by ' '
	from _POSITION
	where num > 2
;	
quit;



*Ce marceau de code est obsolète pour la nouvelle version de la macro;

*------------------------------------------------
Si la variaBle produit est numérique
alors on la transforme en charactère 
en rajoutant la lettre P à chaque modalité
de la variables produit
;
/*
%macro changement_type_var_produit();

%if &type_variable_produit = Num %then %do;
	data &tableEntree;
		drop &product;
		rename prodC=&product;
		set &tableEntree;
		prodC = 'P'||strip(&product);
	run;
%end;

%mend;

%changement_type_var_produit();
*/

data _POSITION;
	set _POSITION;
	paraC=scan(variable, 1, '_');
	timeC=scan(variable, 2,'_');
	if num > 2 ;
	if length(variable)>1;	*Evite le problème d import des colonnes Excel vides;
run;

proc sort data=_POSITION out=_PARA;
	by paraC ;
run; 
	
data _PARA;
	set _PARA;
	by paraC ;
	if first.paraC;
run;

proc sort data=_PARA out=_PARA;
	by num;
run; 

data _PARA;
	set _PARA;
	paraN = _n_;
run;

proc sort data=_POSITION out=_TIME nodupkey;
	by timeC;
run; 

proc sort data=_TIME out=_TIME;
	by num;
run; 

data _TIME;
	set _TIME;
	timeN = _n_;
run;

proc sql;
	select paraC, paraN
	into : listeParaC separated by ' ', : listeParaN separated by ' '
	from _PARA
	order by paraN
;
	select paraC
	into : listeFormatPara separated by '!' 
	from _PARA
	order by paraN
;
	select timeC, timeN
	into : listeTimeC separated by ' ', : listeTimeN separated by ' '
	from _TIME
;	
quit; 

/*
data _PROD;
	set &tableEntree;
	prodC = compress(&product);
	obs=_n_;
run;

proc sort data=_PROD out=_PROD nodupkey;
	by prodC;
*/

data _PROD;
	set &tableEntree;
	obs=_n_;
run;
proc sort data=_PROD out=_PROD nodupkey;
	by &product;
proc sort data=_PROD out=_PROD;
	by obs;
run;

*
Le nom des produits est recodé pour gérer 
les noms de produits avec des charactères spéciaux
;

data _PROD;
	retain PROD_RECODE;
	set _PROD;
	PROD_RECODE=compbl(cat('P',_n_));
;
run;

proc sql;
	select PROD_RECODE
	into : listeProdC separated by ' '
	from _PROD
	order by obs
;
	select &product
	into : listeFormatProd separated by '!'
	from _PROD
	order by obs
;
quit;


proc sql; 
	create table &tableEntree as
	select a.PROD_RECODE, b.*
	from _PROD as a, &tableEntree as b
	where a.&product=B.&product
	;
quit;

data &tableEntree;
	rename PROD_RECODE=&product;
	drop &product;
	set &tableEntree;
run;


%let	variablesDeClassement= &varClassement;
%let	variablesTransposees= &varTransposees;
%let	listeParametres = &listeParaC ; 					
%let	listeFormatsParametre = &listeFormatPara ; 				
%let	listeProduits = &listeProdC ;						
%let	listeFormatsProduit = &listeFormatProd;				
%let	listeTemps = &listeTimeC ;							
%let	listeTempsNum = &listeTimeN;						
 

%put _global_;




*-----------------------------------------------
En cas de variable importer au format charactère
On s'assure que toutes les variables sont bien 
numériques
;

proc sort data=&tableEntree out=&tableEntree;
	by &varClassement;
run;
proc transpose data=&tableEntree out=&tableEntree;
	var &varTransposees;
	by  &varClassement;
run;

data &tableEntree;
	set &tableEntree;
	value=input(col1, numx32.);
run;

proc transpose data=&tableEntree out=&tableEntree;
	var value;
	id _name_;
	by &varClassement;
run;







%macro miseEnFormeDesDonnees();


*-----------------------------------------------------------------------
1. Macro-variable avec le nombre de parametres, de produits et de temps

------------------------------------------------------------------------; 

%global NombreParametres NombreProduits NombreTemps;

%let NombreParametres 	= %sysfunc(countw(&listeParametres,%str( )));
%let NombreProduits 	= %sysfunc(countw(&listeProduits,%str( )));
%let NombreTemps 		= %sysfunc(countw(&listeTemps,%str( )));



*-------------------------------------------------------------------------
2. Stockage dans des macro-variables :
	- chaque parametre (para1, para2...) et numérique (ParaN1, ParaN2)
	- chaque produit en charactère (p1, p2...) et (pn1, pn2)   
	- chaque temps en charactère (t1, t2...) et (tn1, tn2)
--------------------------------------------------------------------------; 

%macro StockageVariables ();  

	%do i = 1 %to &NombreParametres;			
		%global para&i paraN&i; 									


		%let Para&i 	= %scan (&listeParametres, &i, ' ' ) ;
		%let paraN&i 	= &i;

	%end;



	%do i = 1 %to &NombreProduits;				
		%global p&i pN&i; 								
		%let p&i 	= %scan (&listeProduits, &i, ' ' ) ;
		%let pN&i 	= &i;
	%end;

	%do i = 1 %to &NombreTemps;							
		%global t&i tN&i; 
		%let t&i = %scan (&listeTemps, &i, ' ' ) ;
		%let tN&i = %scan (&listeTempsNum, &i, ' ' ) ;
	%end;

%mend StockageVariables;


%StockageVariables();



*--------------------------------------------------------------------------------------
3. Création des formats paramétre (parameterf.), product (productf.) et temps (tempsf.)
---------------------------------------------------------------------------------------; 

%macro format();

*==================
Format parametre

===================;

%global paraf1;
%let paraf1 = %scan(&listeFormatsParametre, 1, !|);
%let listeFormatsParametre2 = %str( 1 = "&paraf1" ) ;

%do i = 2 %to &NombreParametres;
	%global paraf&i;
	%let paraf&i = %scan(&listeFormatsParametre, &i,!); 

	%let listeFormatsParametre2 = &listeFormatsParametre2 %str( &i = "&&Paraf&i" ) ;
%end;





*==================
Format produit
===================;


%global pf1 listeInformatsProduit;
%let pf1 = %scan(&listeFormatsProduit, 1,!);
%let listeFormatsProduit2 = %str( 1 = "&pf1" ) ;
%let listeInformatsProduit 	= %str ( "&pf1" = 1 ) ; 

%do i = 2 %to &NombreProduits;
	%global pf&i;
	%let pf&i = %scan(&listeFormatsProduit, &i, !); 
	%let listeFormatsProduit2 = &listeFormatsProduit2 %str( &i = "&&pf&i" ) ;
	%let listeInFormatsProduit = &listeInFormatsProduit %str( "&&pf&i" = &i ) ;
%end;

*Format pour la comparaison des produits 2 à 2;
%do i=1 %to %eval(&NombreProduits-1);

%do j=%eval(&i+1) %to &NombreProduits;

		*-Création de la variable numérique comparaison des produits
		1000 = comparaison des produits
		100 * i + j = produit i comparé au produit j
		;
		%let numTimeN = %eval(1000+100*&i+&j);
		%let numDuformat= %eval(&NombreProduits+&i);


		%let pf&numDuFormat= &&pf&i vs &&pf&j;
		%let listeFormatsProduit2 = &listeFormatsProduit2  %str( &numTimeN = "&&pf&numDuFormat" );
		%let listeInFormatsProduit = &listeInFormatsProduit %str("&&pf&i vs &&pf&j" = &numTimeN);
	%end;

%end;

*=====================================
			Format temps
=====================================;


*Initialisation liste des formats pour le temps;
%let listeFormatsTemps 		=	 %str( &tN1 = "&t1" ) ;

*Initialisation liste des informats pour le temps;
%global listeInformatsTemps;
%let listeInformatsTemps 	= %str ( "&t1" = &tN1 ) ; 

%do i=2 %to &NombreTemps;

	%let listeFormatsTemps = &listeFormatsTemps %str( &&tN&i = "&&t&i" ) ;
	%let code_Numerique_ti_t1 = %sysevalf(1000+&&tN&i);
	%let listeFormatsTemps = &listeFormatsTemps %str( &code_Numerique_ti_t1 = "&&t&i.-&t1" ) ;
	%let listeInformatsTemps = &listeInformatsTemps %str( "&&t&i" = &&tN&i ) ;
	%let code_Charactere_ti_t1 = &&t&i-&t1;
	%let listeInformatsTemps = &listeInformatsTemps %str( "&code_Charactere_ti_t1" = &code_Numerique_ti_t1 ) ;

%end;







%if %length(&libFormat) ne 0 %then %do;
	%let catalogue_format = &libFormat;
%end;

%else %do;

	%let catalogue_format = work;

%end;

proc format library = &catalogue_format;

		value parameterf
			&listeFormatsParametre2;

		value productf
			&listeFormatsProduit2;

		invalue productf
			&listeInformatsProduit;

		invalue timef
			&listeInformatsTemps;

		value timef
			&listeFormatsTemps;
run;

%mend format; 

%format();



*-------------------------------------------------------------------------
4. Transposition de la table entrée en long
--------------------------------------------------------------------------; 

proc sort data=&tableEntree 
	out=&tableSortie; 
	by &autresVariables &variablesDeClassement;

proc transpose data=&tableSortie 


	out=&tableSortie (rename=(col1=value)) name=variable; 
	var &variablesTransposees;
	by &autresVariables &variablesDeClassement;
	informat col1 best32.;

	format col1 best32.;

run;



*
Supression des espaces éventuels dans les noms de produits
CJU le 27/07/2017

;

data &tableSortie;
	retain subj prod variable value;

	rename prod=product subj=subject;
	set &tableSortie;
	prod=compress(&product);
	subj=&subject;
run;





*-------------------------------------------------------------------------
5. Création des variables numériques :
	- paramètres caractère et numérique 
	- produit numérique 
	- temps caractére
--------------------------------------------------------------------------; 

data &tableSortie;

	informat 	parameter $char100.
				product $char100. 
				time $char100.
				productN 8. 
				parameterN 8. 
				timeN 8.

	; 

	format		parameter $100.
				product $char100.
				time $100.
				parameterN parameterf.
				productN productf.
				timen timef.
	;

	label 		subject='Subject'	
				parameter='Parameter'	
				parameterN='Parameter'
				product='Product'
 				productN='Product'
				value='Value'

				variable='Variable entrée'

	;



	set &tableSortie;



		%do i = 1 %to &NombreTemps;
			if index(upcase(variable), %upcase("&&t&i")) ne 0 then do; 
				time = "&&t&i";  

				timeN= "&&tN&i";

			end;	

		%end;

		%do i = 1 %to &NombreParametres;
			*----------------------------------------
			CJU le 23/11/2017
			Correction erreur quand 2 paramètres 
			avait un nom trop proche

			;

			if upcase(scan(variable,1,'_')) = %upcase("&&Para&i") then do; 
			*if index(upcase(variable), %upcase("&&Para&i")) ne 0 then do; 
				parameter = "&&Para&i"; 
				parametern=	&i; 
			end;	
		%end;

		%do i = 1 %to &NombreProduits;
			if (upcase(product) = %upcase("&&p&i")) ne 0 then do; 
				productn = &i; 
			end;

		%end;

run;

*========================================================================
Vérification de la bonne correspondance des nouvelles variables

=========================================================================;



proc sql;

	select distinct put(parametern,8.0) as num, parametern, parameter,variable
	from &tableSortie;


	select distinct productn,product
	from &tableSortie;



	select distinct time,variable

	from &tableSortie;

quit;





*============================================================================
Calcul des variations par rapport à la valeur basale Ti-T0 pour chaque sujet
=============================================================================;

%if &delta = abs or &delta = rel %then %do;


		proc sort data=&tableSortie out=&tableSortie; 
			by &autresVariables subject parametern parameter productn product timen time; 
		run;
		
		proc transpose data=&tableSortie out=&tableSortie;
			var value;
			id time;
			by &autresVariables subject parametern parameter productn product ;
		run;
	
		data &tableSortie;
			set &tableSortie;
			%do i = 2 %to &NombreTemps;
				%if &delta=abs %then %do; &&t&i.._&t1 = &&t&i-&t1; %end;
				%if &delta=rel %then %do; 

					&&t&i.._&t1 = (&&t&i-&t1)/&t1; 

					%put ERROR- %str(ATTENTION : Analyse sur les variations relatives !); 
				%end;

			%end;



		proc transpose data=&tableSortie out=&tableSortie name=time;

			var &t1--&&t&NombreTemps.._&t1;

			by &autresVariables subject parametern parameter productn product;
		run;

%end;






*========================================================================
	Calcul des différences entre les produits pour chaque sujet 
=========================================================================;

%if ((&plan_experimental = Intra_individuel) and (&NombreProduits >1)) %then %do;

		data &tableSortie;
			set &tableSortie;
		run;

		proc sort data=&tableSortie out=&tableSortie; 
			by &autresVariables subject parametern parameter  time;

		proc transpose data=&tableSortie out=&tableSortie name=time;
			var value;

			id product;
			by &autresVariables subject parametern parameter  time;
		run;

		data &tableSortie;
			set &tableSortie;
			%do i=1 %to %eval(&NombreProduits-1);
				%do j=%eval(&i+1) %to &NombreProduits;
					&&p&i.._&&p&j = &&p&i - &&p&j;
				%end;
			%end;

		run; 

		%let nombreProduits_1 = %eval(&NombreProduits-1);



		proc transpose data=&tableSortie out=&tableSortie (rename=(col1=value)) name=time name=product;
			var &p1--&&p&nombreProduits_1.._&&p&NombreProduits;
			by &autresVariables subject parametern parameter time;
		run;

%end;


data &tableSortie;



	retain &autresVariables Subject Parameter ParameterN ProductN Product TimeN Time value;
	
	rename 	Parameter=parameterC 


			parameterN=parameter 

			product=productC
			productN=product

			Time=timeC 
			TimeN=time 
			;


	label	
			timeN='Time'  
			productN ='Product'
			parameterN ='Parameter'	
	;

	informat 	time $char100. 
				timen 8. 
				value best32.;

	format 	time $char100. 
			timen timef.
			productN productf.
			value best32.;


	length product $ 100;



	set &tableSortie;


		%do i = 1 %to &NombreTemps;
			if index(upcase(time), %upcase("&&t&i")) ne 0 then timeN = &&tN&i;			


		%end;
		

		%do i = 2 %to &NombreTemps;



			if index(upcase(time), %upcase("&&t&i.._&t1")) ne 0 then timeN = %sysevalf(&&tN&i+1000);
		*!!! if (  index(upcase(time), %upcase("&&t&i.._&t1")) ne 0  and  index(upcase(product), %upcase("&p1._&p2")) ne 0 ) then timeN = %sysevalf(&&tN&i+2000);

		%end;



		%if (&plan_experimental = Intra_individuel) %then %do; 

			%do i = 1 %to &NombreProduits;
				if (upcase(product) = %upcase("&&p&i")) ne 0 then do; 
					productn = &i; 
				end;
			%end;

			%do i=1 %to %eval(&NombreProduits-1);
				%do j=%eval(&i+1) %to &NombreProduits;
					if (upcase(product) = %upcase("&&p&i.._&&p&j")) ne 0 then productn=%sysevalf(1000+100*&i+&j);
				%end;

			%end;



		%end;

run;


*================================================================



Vérification de la bonne correspondance des nouvelles variables

+ effectif à chaque temps par produit et paramètre
=================================================================;

proc sql;

	select distinct time, timeC
	from &tableSortie;

	select distinct put(parameter,8.0) as num, parameter, product, time, count(subject) as Effectif
	from &tableSortie
	group by parameter, product, time;
quit;


*================================================================

Suppression des variable de type caractère

=================================================================;


data &tableSortie;
	drop parameterC productC timeC;	
	set  &tableSortie;
	NO=_n_; 
run;

*================================================================
Propriété des variables
=================================================================;

proc contents data=&tableSortie;
ods select variables;

run;




%put _global_;



%if %length (&tableVar) ne 0 %then %do;

	data &tableVar;  			 
   		set sashelp.vmacro(where=(scope='GLOBAL'));    

		if substr(name,1,3) ne 'SYS';      

	run;
%end;

%mend;

%miseEnFormeDesDonnees();







*
Nettoyage des tables
;
proc datasets library=work;	
delete 	
_position
_para 
_prod 
_time
;
run;



%let 	table=all;
%let 	numParametre= ;	 
%let 	format=8.&Nombre_decimales;
%let 	affichage =	;
%let 	stat = 2.5 7 8; 
%let 	pol = calibri;
%let	tail = 10;
%let	larg = 1;
*%let 	lang = ang;
%let	leg = Y	;


%macro stat(); 

*---------------------------------------
1. Définition des formats
;



proc format;
	invalue statf
		'N_VM' 		= 1
		'M_EC' 		= 2
		'M_ESM'		= 2.5
		'ESM'  		= 3
		'Med'  		= 4
		'Q1_Q3'   	= 5
		'Range'		= 6
		'pvalueC'	= 7
		'Signific'	= 8
	
;

	value statf
		1 	= 'N (miss)'
		2 	= 'mean (SD)'
		2.5 = 'Mean_SEM'	

		3 	= 'SEM'

		4 	= 'median'
		5 	= 'Q1 ; Q3'
		6 	= 'min ; max'
		7 	= 'p_value'
		8	= 'Significant'
	;

value statfr

		1 	= 'N (VM)'
		2 	= 'moy (ET)'
		2.5 = 'Moy_ESM'	
		3 	= 'ESM'

		4 	= 'médiane'

		5 	= 'Q1 ; Q3'
		6 	= 'min ; max'
		7 	= 'p-valeur'
		8	= 'Significatif'
	;
/*
%if %upcase(&lang) = ANG %then %do;
	value statf
		1 	= 'N (miss)'

		2 	= 'mean (SD)'
		2.5 = 'Mean_SEM'	
		3 	= 'SEM'
		4 	= 'median'

		5 	= 'Q1 ; Q3'

		6 	= 'min ; max'
		7 	= 'p-value'
		8	= 'Significant'
	;
%end;

%else %do;
	value statf

		1 	= 'N (VM)'
		2 	= 'moy (ET)'
		2.5 = 'Moy_ESM'	
		3 	= 'ESM'

		4 	= 'médiane'

		5 	= 'Q1 ; Q3'
		6 	= 'min ; max'
		7 	= 'p-valeur'
		8	= 'Significatif'
	;
%end;

	value significativityf

		1	=	'Yes'
		0	=	'No'
;

	value $testf
		'Signed Rank'	=	'Wilcoxon'
		"Student's t"	=	'Student'
;*/
run;





*-----------------------------------------------------------------
On supprime les sorties des
procédures 
;
ods rtf exclude all;











*----------------------------------------------------------------
Analyse du paramètre sélectionné si numParametre a une valeur 



		sinon analyse de tous les paramètres

;

%if %length(&numParametre) ^= 0 %then %do;
	%let __selection_parametres =%str(if parameter=&numParametre);

%end;
%else %do;
	%let __selection_parametres =%str() ;
%end;



************************************************************


2. CREATION D UNE TABLE AVEC LES STATISTIQUES DESCRIPTIVES



************************************************************;

data __&table;
	set &table;
	&__selection_parametres;
run;


ods output table=__DESC;

proc tabulate data=__&table order=unformatted;

	where ( time < 1000 and product < 1000 ) = 1 or (time > 1000);
	class parameter time  product;
	var value;
	table parameter=' '*time=' '*value=' '*(n nmiss mean median stderr std q1 q3 min max), product ;

run; 







*--------Mise en forme de la table------------------------------------------------;

data __DESC;

	set __DESC;

	N_VM = strip(put(value_N,8.0))||' ('||strip(put(value_nmiss,8.0))||')';

	M_EC = strip(put(value_mean,&format))||' ('||strip(put(value_std,&format))||')';
	ESM  = strip(put(value_stderr,&format));

	Med  = strip(put(value_median,&format));
	Q1_Q3  	= strip(put(value_q1,&format))|| ' ; '||strip(put(value_q3,&format));
	Range = strip(put(value_min,&format)) || ' ; '||strip(put(value_max,&format));


	/*Format Moy ± SEM*/ 

	M_ESM = strip(put(value_mean,&format))||' ± '||strip(put(value_stderr,&format));

run;




*-------Transposition en long pour impression (PROC REPORT)---------------;

proc sort data=__DESC out=__DESC; 
	by parameter product   time;
run;


proc transpose data=__DESC out=__DESC name=stat; 
	by parameter product  time;
	var N_VM--M_ESM;
run;




*--------------------------------------------------------------------------
ENREGISTREMENT DES TESTS STATISTIQUES POUR LES ANNEXES DU RAPPORT
;
ods documents Name=ANNEXE(Write);

****************************************************************************
	2. TEST T DONNES APPARIES / WILXOXON / SHAPIRO-WILK
;

proc sort data=__&table out=__&table; by parameter product time; run;




%if %upcase(&sortieBrute) = N %then %do; ods rtf exclude all; %end;



%else %do; ods rtf select all; %end;


ods output testsfornormality=__NORM testsforlocation=__LOC;

proc univariate data=__&table normal;
where time > 1000 ; 
ods exclude Moments Quantiles ExtremeObs MissingValues ParameterEstimates GoodnessOfFit FitQuantiles;
  var value;
  by parameter product time;
run;



*
CJU
;


*----Table avec choix du test statistique en fonction du Shapiro-Wilk-------;

proc sql noprint;

	create table __ANNEXE1 as

	select  a.parameter, 

			a.product, 
			a.time,
			a.product,
			a.testlab, 
			b.testlab, 

			a.pvalue as shapiro, 
			b.pvalue as pvalue,
			b.test,

			calculated pvalueC as stat,
				case 
					when a.pvalue > 0.01 and b.testlab='t' then 1 
					when a.pvalue <= 0.01 and b.testlab='t' then 0 
					when a.pvalue > 0.01 and b.testlab='S' then 0 
					when a.pvalue <= 0.01 and b.testlab='S' then 1 
					else 0
				end as bool,

				case 

					when b.testlab='t' and b.pvalue ne . then strip(put(b.pvalue,pvalue8.4)||'°')

					when b.testlab='S' and b.pvalue ne . then strip(put(b.pvalue,pvalue8.4)||'*')
					else 'na'
				end as pvalueC



	from __NORM as a, __LOC as b

	where (a.parameter = b.parameter) 
			and (a.product=b.product) 
			and (a.time=b.time) 

;
quit; 


proc sort data=__ANNEXE1  out=__ANNEXE1 ; 
		where test ne 'Sign' and testlab='W';
		by parameter product time shapiro test;
run;
	
proc transpose data=__ANNEXE1 out=__ANNEXE1; 
	var pvalue;
	id test;
	by parameter product time shapiro;
run;



*----Table avec choix du test statistique en fonction du Shapiro-Wilk-------;

proc sql noprint;
	create table __COMP as
	select  a.parameter, 

			a.product, 
			a.time,
			a.product,
			a.testlab, 
			b.testlab, 
			a.pvalue as shapiro, 

			b.pvalue as pvalue,

			b.test,


			calculated pvalueC as stat,

				case 
					when a.pvalue > 0.01 and b.testlab='t' then 1 

					when a.pvalue <= 0.01 and b.testlab='t' then 0 
					when a.pvalue > 0.01 and b.testlab='S' then 0 
					when a.pvalue <= 0.01 and b.testlab='S' then 1 
					else 0

				end as bool,


				case 
					when b.testlab='t' and b.pvalue ne . then strip(put(b.pvalue,pvalue8.4)||'°')
					when b.testlab='S' and b.pvalue ne . then strip(put(b.pvalue,pvalue8.4)||'*')
					else 'na'
				end as pvalueC,



				case

					when b.pvalue = . then 9999
					when b.pvalue < 0.05 then 1
					when b.pvalue >= 0.05 then 0
				end as Signific			

	from __NORM as a, __LOC as b

	where (a.parameter = b.parameter) 
			and (a.product=b.product) 
			and (a.time=b.time) 

			and calculated bool=1
;
quit; 


*---------------------------------
Calcul de la frèquence des test-t
et de Wilcoxon
;



proc sort data=__COMP out=__COMP; 

		where testlab='W'; by parameter product time;

proc transpose data=__COMP out=__PVALUE 
	(rename=(col1=pvalue )) name=stat;
	var pvalue;
	by parameter product time test;
run;

proc transpose data=__COMP out=__COMP  name=stat;
	var pvalueC Signific;
	by parameter product time ;
run;








*--------Création de la variable numérique produit-------------------------------------;


data __RESULT;


	set __DESC __COMP;

	statN = input(stat,statf.);

	format statN statf. time timef.;

		%do a = 1 %to &NombreProduits;
				if index(upcase(product), %upcase("&&p&a")) ne 0 then do; 
					product = &a; 
				end;

		%end;

		%do a=1 %to %eval(&NombreProduits-1);

			%do b=%eval(&a+1) %to &NombreProduits;

				if index(upcase(product), %upcase("&&p&a.._&&p&b")) ne 0 then product=%sysevalf(1000+100*&a+&b);
			%end;

		%end;

run;


*-----------Ajout de la variable pvalue numérique------------------------;

*Récupération des valeurs brutes;

data __RESULT_ti;
	set __RESULT;
	if time < 1000;
run;

*Liaison avec la table contenant les pvalue numérique;
proc sql;
	create table __RESULT_ti_t0 as 

	select *
	from __RESULT as a, __PVALUE as b
	where a.parameter=b.parameter and a.product=b.product and a.time=b.time
	;

quit;



*Concaténation des 2 tables pour récupérer toutes les variables et observation;
data __RESULT;
	set __RESULT_ti __RESULT_ti_t0;
run;








**************************************************************************
	3. CAS OU LE PLAN EXPERIMENTAL = GROUPES PARALLELES
;

%if  &plan_experimental = Groupes_paralleles  %then %do;

	%let comp=0;

	%do produit1=1 %to %eval(&nombreproduits-1);

		%do produit2=%eval(&produit1+1) %to &nombreproduits; 	

	
		%let comp=%eval(&comp+1);

*------------------------------------------------------------------------
			test t pour échantillons indépendants
;

	proc sort data=__&table out=__&table; by parameter time product; 

	ods output statistics=__stat_para&comp 
				ttests=__test_para&comp 
				equality=__equa_para&comp;

	proc ttest data=__&table;
		where time > 1000 and product in (&produit1 &produit2);

		var value;

		class product ;
		by parameter time;
	run;



	proc sql;

		create table __Unpairedttest&comp as
		select a.parameter, a.time, a.class, a.mean, a.stddev, a.stderr, b.variances, b.probt
		from __stat_para&comp as a, __test_para&comp as b


		where (a.parameter = b.parameter) and (a.time=b.time)and class='Diff (1-2)'

	;

		create table __UnpairedTtest&comp as

		select a.parameter, a.time, a.class, a.mean, a.stddev, a.stderr, a.variances, a.probt, b.probf,
		case	
			when b.probf < 0.05 and a.variances='Equal' 	then 0
			when b.probf >= 0.05 and a.variances='Equal' 	then 1  
			when b.probf < 0.05 and a.variances='Unequal' 	then 1
			when b.probf >= 0.05 and a.variances='Unequal' 	then 0
		end as bool


		from __UnpairedTtest&comp as a, __equa_para&comp as b
		where (a.parameter = b.parameter) and (a.time=b.time) and calculated bool=1

	;



	quit;





*------------------------------------------------------------------------
			test de Mann-Whitnney
;

	ods output wilcoxontest=__MannWhitney&comp;
	proc npar1way data=__&table wilcoxon;

		where (time > 1000) and product in (&produit1 &produit2) ;

		var value;

		class product ;
		by parameter time;
	run;



	data __MannWhitney&comp;
		rename nvalue1=MannWhitney;
		keep parameter time label1 nValue1;
		set __MannWhitney&comp;
		if name1 ='P2_WIL';

	run;

	
	proc sql;
		create table __comp_para&comp as
		select *
		from __UnpairedTtest&comp as a, __MannWhitney&comp as b
		where (a.parameter = b.parameter) and (a.time=b.time) ;
	quit;


	proc sql noprint;
		select distinct product
		into: prod1-:prod2
		from __norm
		where product in (&produit1 &produit2) 
		order by product
	;
	quit; 


	proc sort data=__norm out=__norm; by parameter time product; 


	proc transpose data=__norm out=__norm_&comp ;

	where testlab='W';
		by  parameter time;

		var pvalue;
		id product;

	run;

	data __norm_&comp;
		set __norm_&comp;

		if "&prod1"n > 0.01  and "&prod2"n > 0.01 then normalite=1;

		if "&prod1"n <= 0.01  or "&prod2"n <= 0.01 then normalite=0;

	*	time = input(timeTemp, timef.);

	run;

proc sql;

		create table __comp_para&comp as

		select a.*, put(stderr,&format) as stderrC, b.normalite, strip(put(probt,pvalue8.4)||'µ') as ttest, strip(put(MannWhitney,pvalue8.4)||'§') as MW, 

		case

			when b.normalite=1 then calculated ttest
			when b.normalite=0 then calculated MW
		end as pvalueC,

		case

			when b.normalite=1 then probt

			when b.normalite=0 then MannWhitney
		end as pvalue,

		case
			when probt = . then 9999 
			when (b.normalite=1 and probt < 0.05) then 1
			when (b.normalite=1 and probt >= 0.05) then 0

			when MannWhitney = . then 9999 
			when (b.normalite=0 and MannWhitney < 0.05) then 1
			when (b.normalite=0 and MannWhitney >= 0.05) then 0
		end as Signific,

		strip(put(Mean, &format))||' ± '||strip(put(stdDev, &format)) as meanSD,

		strip(put(Mean, &format))||' ± '||strip(put(stderr, &format)) as meanSEM

		from __comp_para&comp as a, __norm_&comp as b

		where (a.parameter = b.parameter) and (a.time=b.time);

	quit;
	

	proc transpose data=__comp_para&comp out=__comp_para&comp;

		by parameter time pvalue ;

		var meanSD meanSEM pvalueC Signific;

	run;


	data __comp_para&comp;

		set __comp_para&comp;
		if _name_ ='meanSD' 	then statN = 2;
		if _name_ ='meanSEM' 	then statN = 2.5;



		if _name_ ='stderrC' 	then statN = 3;

		if _name_ = 'pvalueC' 	then statN = 7;

		if _name_ = 'Signific' 	then statN = 8;
		product = %eval(1000 + 100*&produit1 + &produit2);
	run;

	data __Result;

		set __Result __comp_Para&comp;
	
	run;

	%end;
%end;

%end;

ods document close;

*
IMPRESSION DES RESULTATS

;

*
1 : données brutes

2 : variations (ti-t0)


3 : comparaison des produits sur (ti-t0)

4 :  
;





proc sort data=__RESULT out=__RESULT; by parameter  product time statN; run;


%if %upcase(&TableauRapport) = N %then %do; 
	ods rtf exclude all; %end;
%else %do;
	ods rtf select all; %end;





/*
title1 f=arial h=10pt j=left "The table below presents the change from baseline for the %lowcase(&&paraf&numParametre), for each time point by product";

title2 f=arial bold underlin=1 h=10pt j=center "Table xxx Change from baseline – %lowcase(&&paraf&numParametre)";

*/



proc format;
	value color
		1="#5E64A4"
		2="#D9D9EA"
	;
run;


%if %upcase(&lang) = ANG %then %do; 

	%let Parameter_label=Parameter;
	%let Product_label=Product;
	%let Kinetic_label=Kinetic;	
	%let Mean_label = Mean ± SEM;
	%let pvalue_label = p-value;
	%let Signific_label = Statistically/significant;
	%let test_para_intra = %str(°paired t-test);
	%let test_non_para_intra = %str(*Wilcoxon signed rank test);		
	%let test_para_inter = %str(µUnpaired t-test);
	%let test_non_para_inter = %str(§Mann-Whitney test);
	%let statistical_analysis = STATISTICAL ANALYSIS;
	%let Statistical_comparisons = Statistical comparisons;
	%let study = STUDY;
    %let TT1 = 1. Statistical comparisons;
	%let TT2 = 2.Descriptive statistics;

	proc format;

		value signific

			1 = 'Yes'

			0 = 'No'

			9999 = 'na';



	value significant

		1 = 'red'
		0,9999 = 'black'
	run;





%end;
%else %do;
	%let Parameter_label=Paramètre;
	%let Product_label=Produit;

	%let Kinetic_label=Temps;
	%let Mean_label = Moy ± ESM;

	%let pvalue_label = p-valeur;

	%let Signific_label = Statistiquement/significatif;

	%let test_para_intra = %str(test-t pour échantillons appariés);

	%let test_non_para_intra = %str(*test des rangs signés de Wilcoxon);
	%let test_para_inter = %str(µt-test pour échantillons indépendants);

	%let test_non_para_inter = %str(§test de Mann-Whitney);
	%let statistical_analysis = ANALYSE STATISTIQUE;
	%let Statistical_comparisons = Comparaisons statistiques;
	%let study = ETUDE;

    %let TT1 = 1. Comparaisons statistiques;
	%let TT2 = 2. Statistiques descriptives;

	proc format;
	value signific
			1 = 'Oui'
			0 = 'Non'
			9999 = 'na';

	value significant
		1 = 'red'

		0,9999 = 'black'

	run;

%end;



data _null_;
	x=put(today(),ddmmyy10.);

*	x='* '||put(today(),mmddyy10.)||' *';

	call symput('date',x);

run;


%let couleur_bordures	=	#5E64A4;

%let couleur_fond 		=	#D9D9EA;

proc sort data=__RESULT out=__RESULT2; by parameter product time; run;

proc transpose data=__RESULT2 out=__RESULT2; 
*Pour avoir Mean (SD) : where statN in (1 2 7 8 ) and time > 1000;
	where statN in (&stat) and time > 1000;

	by parameter  product time pvalue;
	id statN;

	var col1;
run;


data __Result2;

	drop significant;

	rename significantN=significant;

	set __Result2;
	significantN=input(significant,best32.);
run;


 
 /*mgu start update methode fr vs ang*/
ODS escapechar='^';

title1 h=14pt f=calibri bold "&study : &num_etude";
title2 " "; 
title3 h=14pt f=calibri bold "&statistical_analysis"; 
title4 " "; 
title5 h=14pt f=calibri bold "&date"; 
title6  " "; 

%if (&plan_experimental=Intra_individuel and %upcase(&lang) = ANG ) %then %do;
title7 h=10pt f=calibri j=left "Statistical software:^2n
 SAS v9.4 ^2n 
 Methodology:^2n
 For each statistical comparison, if the normality assumption was not rejected using a Shapiro-Wilk test (^R/RTF'\u945\a'=0.01), a paired t-test was performed.
 In case of normality rejection, a non-parametric approach was carried out using a Wilcoxon signed rank test.^2n
 The Statistical tests are two-tailed.^2n
 The type I error is set at ^R/RTF'\u945\a'=0.05.^2n";
%end;


%if (&plan_experimental=Intra_individuel and %upcase(&lang) ne ANG ) %then %do;
title7 h=10pt f=calibri j=left "Logiciel statistique:^2n
 SAS v9.4 ^2n 
 Méthodologie:^2n
 Pour chaque comparaison statistique, si l'hypothèse de normalité n'a pas été rejetée à l'aide d'un test de Shapiro-Wilk (^R/RTF'\u945\a'=0.01), un test t apparié a été réalisé.
 En cas de rejet de normalité, une approche non paramétrique a été réalisée à l'aide d'un test de rang signé de Wilcoxon.^2n
 Les tests statistiques sont bilatéraux.^2n
 L'erreur de type I est fixée à ^R/RTF'\u945\a'=0.05.^2n";
%end;


%if (&plan_experimental=Groupes_paralleles and %upcase(&lang)= ANG)%then %do;
title7 h=10pt f=calibri j=left "Statistical software:^2n 
	SAS v9.4^2n 
	Methodology:^2n 
	for each intra-individual comparison, if the normality assumption was not rejected using a Shapiro-Wilk test (^R/RTF'\u945\a'=0.01), a paired t-test was performed.
 	In case of normality rejection, a non-parametric approach was carried out using a Wilcoxon signed rank test.^2n
 	For each between groups comparison an unpaired t-test was carried out (or Mann-Whitney test according to normality check).^2n	
 	The Statistical tests are two-tailed.^2n
 	The type I error is set at ^R/RTF'\u945\a'=0.05.^2n";
%end;


%if (&plan_experimental=Groupes_paralleles and %upcase(&lang) ne ANG)%then %do;
title7 h=10pt f=calibri j=left "Logiciel statistique:^2n 
	SAS v9.4^2n 
	Méthodologie:^2n 
	Pour chaque comparaison intra-individuelle, si l'hypothèse de normalité n'a pas été rejetée à l'aide d'un test de Shapiro-Wilk (^R/RTF'\u945\a'=0.01), un test t apparié a été réalisé.
 	En cas de rejet de normalité, une approche non paramétrique a été réalisée à l'aide d'un test de rang signé de Wilcoxon.^2n
 	Pour chaque comparaison entre les groupes, un test t non apparié a été effectué (ou test de Mann-Whitney selon le contrôle de normalité).^2n	
 	Les tests statistiques sont bilatéraux.^2n
 	L'erreur de type I est fixée à ^R/RTF'\u945\a'=0.05.^2n";
%end;

/*************mgu update methode fr vs ang end**************/

ods rtf startpage=now;

ods rtf text="^S={indent=3.5 fontsize=10 pt font_face=calibri font_weight=bold} {\tc\f3\fs0\cf8  &TT1.}";
title8 h=10pt f=calibri bold j=left "&TT1."; 


proc report data=__RESULT2 out=__REPORT_2

		style(report)= [bordercolor=&couleur_bordures frame=box rules=all cellspacing=0 font_face=&pol fontsize=&tail pt ]
		style(header)= [bordercolor=&couleur_bordures BACKGROUND=&couleur_fond font_weight=bold fontsize=&tail pt font_face=&pol cellspacing=0]
		style(column)= [bordercolor=&couleur_bordures cellwidth=&larg.in  fontsize=&tail pt font_face=&pol]
;

		column (/*"&&paraf&numParametre"*/ parameter product time  'Mean_SEM'n pvalue 'p_value'n Significant);
		
		define parameter 	/ 	"&Parameter_label" order=internal group left style=[BACKGROUND=&couleur_fond font_weight=bold] ;
		define product 		/ 	"&Product_label" order=internal group left f=productf. style=[BACKGROUND=&couleur_fond font_weight=bold];

		define time 		/ 	"&Kinetic_label" order=internal group left style=[BACKGROUND=&couleur_fond font_weight=bold] ;


		define pvalue 		/  center  noprint ;
*		define 'mean (SD)'n /	"Mean (SD)" center display style=[color=$significant. cellwidth=1 in];
		define 'Mean_SEM'n /	"&Mean_label" center display style=[color=$significant. cellwidth=1 in];
		define 'p_value'n 	/ 	"&pvalue_label" center display style=[color=$significant. cellwidth=1 in];
		define Significant	/	"&Signific_label" center display f=signific. style=[color=significant. cellwidth=1 in] ;

	
		compute parameter;
     	 	if parameter ^= ''  then call define(_row_,'style','style=[bordertopcolor=&couleur_bordures bordertopwidth=1]');

   		endcomp;



		compute product;

			if product ^= ''  then call define(_row_,'style','style=[bordertopcolor=#46B7C3 bordertopwidth=0.5]');
			if statn=2 then call define (_col_, "style", "style=[fontweight=bold]");



			if statn=7 then call define (_row_, "style", "style=[backgroundcolor=#DAEFF2 fontstyle=italic]");

		endcomp;


	%if &leg = Y %then %do;

		compute after _page_ /left style={font_face=arial fontsize=8pt color=blue bordertopcolor=&couleur_bordures bordertopwidth=1};

			%if ( &plan_experimental = Intra_individuel  and &affichage ne 1 ) %then %do;
				line "&test_para_intra";
				line "&test_non_para_intra";
				line "Note: p-value < 0.05 is statistically significant";
			%end;

			%else %if ( &plan_experimental = Groupes_paralleles and &affichage ne 1) %then %do;
				line "&test_para_intra";
				line "&test_non_para_intra";
				line "&test_para_inter";
				line "&test_non_para_inter";
				line "Note: p-value < 0.05 is considered as statistically significant";
			%end;

		endcomp;
	
	%end;

run;

title;

ods rtf startpage=now;

ods rtf text="^S={indent=3.5 fontsize=10 pt font_face=calibri font_weight=bold} {\tc\f3\fs0\cf8  &TT2.}";

title1 h=10pt f=calibri bold j=left "&TT2."; 

proc report data=__RESULT
		style(report)= [bordercolor=&couleur_bordures frame=box rules=all cellspacing=0 font_face=&pol fontsize=&tail pt ]
		style(header)= [bordercolor=&couleur_bordures BACKGROUND=&couleur_fond font_weight=bold fontsize=&tail pt font_face=&pol cellspacing=0]
		style(column)= [bordercolor=&couleur_bordures cellwidth=0.8 in  fontsize=&tail pt font_face=&pol]
;

%if %upcase(&lang) = ANG %then %do;
			format statN statf.;
		%end;

		%else %do;
			format statN statfr.;;
		%end;

		column (/*"&&paraf&numParametre"*/ parameter product statN time, col1 n)
;
		where time < 1000 and statN in (1 2 4 6) and not missing(col1) 
;	
		define parameter 	/ 	/*"&Parameter_label"*/ " "	order=internal group left style=[BACKGROUND=&couleur_fond font_weight=bold] ;
		define product 		/ 	/*"&Product_label"*/ " "	order=internal group left f=productf. style=[BACKGROUND=&couleur_fond font_weight=bold];
		define statN 		/ 	/*"Stat"*/ 			" "	order=internal group left style=[BACKGROUND=&couleur_fond font_weight=bold] ;
		define time 		/ 	/*"&Kinetic_label"*/ " "	across order=internal   style=[BACKGROUND=&couleur_fond font_weight=bold] ;
		define col1			/	" " 		center display ;
		define n			/ 	noprint ;

;
run;



ods rtf startpage=now;
ods rtf text="^S={indent=3.5 fontsize=10 pt font_face=calibri font_weight=bold} {\tc\f3\fs0\cf8 3. Box-plot}"; 
title1 h=10pt f=calibri bold j=left "3. Box-plot";
/*mgu update start*/
ods graphics on / 
      width=10in HEIGHT=6in
      imagefmt=jpeg
      imagemap=on
      imagename="MyBoxplot"
      border=off
;
/*********************/
PROC SGPANEL  DATA=__all;
styleattrs 
		datacolors=(CX1C3583 CXA6A6A6 CXEF7B0B CXA0A1CC CX66FFFF CXCCFF99 CX5E64A4 CXF8B475 CX5274DA)
		datacontrastcolors=(black);/*mgu update color Eurofins*/

PANELBY parameter/uniscale=column columns=1;
  VBOX value / category = time group=product;
  where time<1000 and product<1000;/*mgu update 22/09/2021*/
RUN;
title;
/*
ods rtf text="^S={indent=3.5 fontsize=10 pt font_face=calibri font_weight=bold} {\tc\f3\fs0\cf8 4.SAS output}"; 
title1 h=10pt f=calibri bold j=left "4.SAS output";
*/
%mend;

ods rtf file="d:\Techniciens\Macro_SAS\&num_etude..rtf" notoc_data bodytitle nogtitle startpage=no contents=yes  notoc_data style=sty.Eurofins1;
*
ods rtf text = "^S={outputwidth=100% just=l}{\field{\*\fldinst {\\TOC \\f \\h}}}";  

%let	sortieBrute=N;
%let	TableauRapport=	O;
%stat();

/*ods rtf startpage=now;

title h=12pt f=calibri bold "Appendice 1"; 
proc print data=__ANNEXE1;

	var parameter product time shapiro 'Signed Rank'n 'Student''s t'n;

	format product productf. time timef.;

run;*/

ods rtf startpage=now;

ods rtf text="^S={indent=3.5 fontsize=10 pt font_face=calibri font_weight=bold} {\tc\f3\fs0\cf8 4. SAS output}"; 
title1 h=10pt f=calibri bold j=left "4. SAS output";
proc document name=ANNEXE;
	replay;
run;

/*
%let	sortieBrute=o;
%let	TableauRapport=	N;
options symbolgen mprint mlogic;
%stat();
*/
ods rtf close;











*---------------------------------------------------------------

SUPPRESSION DES SAUTS DE SECTION DANS LA SORTIE RTF FINALE 
;


/*      Reading and writing to the same file at the same time can be trouble depending on where you are.  */ 
   data temp ;    
     length line $10000;
     infile "d:\Techniciens\Macro_SAS\&num_etude..rtf" length=lg lrecl=1000000 end=eof;
     input @1 line $varying10000. lg;
	run;



  data _null_;
      set temp ; retain flag  0 ; 

      file  "d:\Techniciens\Macro_SAS\&num_etude..rtf";
/*  KEEP FIRST SECTION BREAK AND ELIMINATE THE REST */
      if (index(line,"\sect") > 0) then do ;

         if ( flag ) then  

            line = "{\pard\par}" ;

         else do ;

            flag = 1 ;

         end ; 
      end ;

    put line ;  

    run;





	
proc datasets library=temp nolist kill;

quit;

run;



libname d "d:\Statistiques\Data_SAS\Dermscan";

proc sql; 
	insert into d.liste_etudes_essais

set Date = today()  ,	
	heure =time() 	,	
	Etude = "&num_etude",
	Nombre_prod = input(symget('NombreProduits'),best32.),	
	Nb_para	= input(symget('NombreParametres'),best32.),

	Nb_tps	= input(symget('NombreTemps'),best32.),
	Type = "&plan_experimental",
	Langue = "&lang",	
	Nb_dec = "&nombre_decimales&",
	Nb_dec = "&nombre_decimales&",
	listeParametres = "&listeParaC" ,
	listeTemps = "&listeTimeC" ,
	listeProduits = "&listeProdC",
	auteur = "&_CLIENTUSERID",

	site= "&site"
;
quit;

