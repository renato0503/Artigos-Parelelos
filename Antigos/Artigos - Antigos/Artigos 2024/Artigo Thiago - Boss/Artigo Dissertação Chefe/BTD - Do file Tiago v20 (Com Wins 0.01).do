clear all
log using "meu_resultado_AT.smcl", replace

cd  "C:\Users\20192MPC019\Downloads" 
import excel "Base BTD edit.xlsx", sheet("Sheet1") firstrow

save "BDs_tiago.dta", replace

*transformando ausencia de informação em missing

ds, not(type numeric)
destring ID, replace force


*exibir em tabela os setores industriais
tabulate Setor

* Comandos para excluir setores da amostra da pesquisa  
drop if Setor== "FinançaseSeguros"
drop if Setor== "Fundos"

sort ID Ano
duplicates drop ID Ano, force
xtset ID Ano, yearly


*gerando variáveis defasadas

gen lag_AT = AT[_n-1]
gen lag_CR = CR[_n-1]
gen lag_RECEITA = RECEITA[_n-1]

*gerando variáveis para a regressão GR (Jones Modificado)
 
gen AT_lagAT=1/lag_AT
replace RNO = 0 if RNO ==.
gen TA = LL - RNO - FCO
gen Accruals = LL - RNO - FCO
gen GR=Accruals/lag_AT
gen deltaRECEITA=(RECEITA-lag_RECEITA)/lag_RECEITA
gen deltaCR=CR-lag_CR
gen beta1=1/lag_AT
gen beta2=(deltaRECEITA/lag_AT)-(deltaCR/lag_AT)
gen beta3=INV/lag_AT

areg GR beta1 beta2 beta3, absorb(ID)

* winsorizando...
winsor beta1, gen(Wbeta1) p(0.01)
winsor beta2, gen(Wbeta2) p(0.01)
winsor beta3, gen(Wbeta3) p(0.01)
winsor GR, gen(WGR) p(0.01)

areg WGR Wbeta1 Wbeta2 Wbeta3, absorb(ID)

*****predict residuo_GR, dresiduals

reg WGR Wbeta1 Wbeta2 Wbeta3

predict residuo_GR, residuals


*GERANDO VARIÁVEIS


gen dummyConservadorismo = 0
replace dummyConservadorismo = 1 if CONSERVADORISMO<0
winsor CONSERVADORISMO, gen(WCONSERVADORISMO) p(0.01)
gen dummyNOL = 0
replace dummyNOL = 1 if NOL<=0


** winsorizando

winsor BTD_AT, gen(WBTD) p(0.01)
winsor BTD_LN, gen(WBTDLN) p(0.01)
winsor RTA, gen(WRTA) p(0.01)

*winsor residuo_CONSERVADORISMO, gen(Wresiduo_CONSERVADORISMO) p(0.01)
*WINSORIZAR Variáveis do BTD

winsor LNAT, gen(WLNAT) p(0.01)
winsor INV_LN, gen(WINV) p(0.01)
winsor deltaRECEITA, gen(WdeltaRECEITA) p(0.02)
winsor NOL, gen(WNOL) p(0.01)
winsor IRD_LN, gen(WIRD) p(0.01)
winsor PLLL_LN, gen(WPLLL) p(0.01)
winsor AFDPFD_LN, gen(WAFDPFD) p(0.01)

*Gerando variável de interação

gen CGR=CONSERVADORISMO*residuo_GR

*Gerando variável de interação com a dummy conservadorismo

gen dCGR=dummyConservadorismo*residuo_GR

*dummy de ano vai ser n-1 variáveis, sendo 1 o ano da variável, 0 caso contrário.

egen IDAno = group(Ano), label
egen IDSetor = group(Setor), label

* WINSSORIZAR o modelo de ETR
winsor ETR, gen(WETR) p(0.01)
winsor IRCS_LN, gen(WIRCS) p(0.01)
winsor LAIR_LN, gen(WLAIR) p(0.01)
winsor EBIT_LN, gen(WEBIT) p(0.01)
winsor ROA, gen(WROA) p(0.02)
winsor MTB, gen(WMTB) p(0.02)
winsor SALES, gen(WSALES) p(0.01)
winsor residuo_GR, gen(Wresiduo_GR) p(0.01)
winsor dCGR, gen(WdCGR) p(0.02)
winsor CGR, gen(WCGR) p(0.02)


** TABELA GERAL DA ESTISITICA DESCRITIVA

univar WBTD WETR dummyConservadorismo WCONSERVADORISMO Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD WdCGR WCGR WIRCS WLAIR WEBIT WROA WMTB WSALE

* a.1)gerando BTD com interação
*dummyConservadorismo
xtreg WBTD dummyConservadorismo WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD i.IDAno i.IDSetor

xtreg WBTD dummyConservadorismo Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD WdCGR i.IDAno i.IDSetor

sum WBTD dummyConservadorismo Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD WdCGR i.IDAno i.IDSetor

*WCONSERVADORISMO

xtreg WBTD WCONSERVADORISMO Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD WCGR i.IDAno i.IDSetor

sum WBTD WCONSERVADORISMO Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD WCGR i.IDAno i.IDSetor

* b) BTD

xtreg WBTD Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD i.IDAno i.IDSetor

sum WBTD Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD i.IDAno i.IDSetor


* Correlação de perason para todas as variáveis
pwcorr WBTD WETR dummyConservadorismo WCONSERVADORISMO Wresiduo_GR WLNAT WINV WdeltaRECEITA dummyNOL WIRD WPLLL WAFDPFD WdCGR WCGR WIRCS WLAIR WEBIT WROA WMTB WSALES, sig

log close