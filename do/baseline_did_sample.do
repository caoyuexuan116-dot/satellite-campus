version 17.0
clear all
set more off

* Project root and output directory. Stata uses forward slashes reliably on Windows.
local root "C:/Users/Dell/Documents/学习wos数据使用"
local out "`root'/outputs"

cap mkdir "`out'"
log using "`out'/baseline_did_sample.log", replace text

* Import the balanced person-year panel created by scripts/clean_wos_sample.py.
di as text "Importing person-year panel"
import delimited using "`out'/wos_author_year_panel_sample.csv", clear varnames(1) encoding("UTF-8")

compress
destring school_id treated relo_year year post did pub_count cites_wos_core_sum cites_all_db_sum, replace force

* Audit the estimation sample before running regressions.
di as text "Sample checks"
tab school_cn treated, missing
tab year treated, missing
tab treated post, missing
assert inrange(year, 2000, 2022)
assert treated == 0 | treated == 1
assert post == 0 if treated == 0
assert did == treated * post

* Panel identifier is school-person, not just person_id, because an ORCID can
* appear under more than one school in this pilot sample.
egen person_num = group(school_id person_id)
xtset person_num year

label var pub_count "Publication count"
label var cites_wos_core_sum "Times cited, WoS Core"
label var cites_all_db_sum "Times cited, all databases"
label var did "Treated x post"

di as text "Small-sample warning: only 10 school clusters; coefficients are pipeline tests, not final inference."

estimates clear

* Baseline DID: individual fixed effects via xtreg, calendar-year fixed effects,
* and school-clustered standard errors.
xtreg pub_count did i.year, fe vce(cluster school_id)
estimates store m_pub

xtreg cites_wos_core_sum did i.year, fe vce(cluster school_id)
estimates store m_cite_core

xtreg cites_all_db_sum did i.year, fe vce(cluster school_id)
estimates store m_cite_all

estimates table m_pub m_cite_core m_cite_all, b(%9.4f) se(%9.4f) stats(N r2_w)

* Save the imported/checked Stata panel for follow-up diagnostics.
save "`out'/wos_author_year_panel_sample.dta", replace

log close
