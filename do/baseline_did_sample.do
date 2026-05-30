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

* Multi-period DID checks: each treated school should switch on in its own reform year.
tab school_cn relo_year if treated == 1, missing
tab year did if treated == 1, missing

preserve
collapse (mean) treated post did, by(school_cn year relo_year)
list school_cn year relo_year treated post did if treated == 1 & inrange(year, relo_year - 2, relo_year + 2), sepby(school_cn)
restore

* Person-count diagnostics after Python filters outside-school coauthors.
preserve
keep school_cn school_id person_id
duplicates drop
bysort school_cn: gen n_person = _N
bysort school_cn: keep if _n == 1
list school_cn n_person, clean
restore

preserve
collapse (sum) total_pub = pub_count, by(school_cn person_id)
sum total_pub, detail
restore

* Panel identifier is school-person, not just person_id, because an ORCID can
* appear under more than one school in this pilot sample.
egen person_num = group(school_id person_id)
xtset person_num year

* Mark treated-school authors who were already observed before their school opened
* the same-city satellite campus. Controls remain in the robustness sample.
gen pre_period = year < relo_year if treated == 1
replace pre_period = 0 if treated == 0
bysort person_num: egen pre_pub_treated = total(pub_count * pre_period)
gen incumbent_treated_author = 1
replace incumbent_treated_author = 0 if treated == 1 & pre_pub_treated == 0

label var pub_count "Publication count"
label var cites_wos_core_sum "Times cited, WoS Core"
label var cites_all_db_sum "Times cited, all databases"
label var did "Treated x post"
label var incumbent_treated_author "Treated-school author observed before reform"

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

xtreg pub_count did i.year if incumbent_treated_author == 1, fe vce(cluster school_id)
estimates store m_pub_incumbent

estimates table m_pub m_cite_core m_cite_all m_pub_incumbent, b(%9.4f) se(%9.4f) stats(N r2_w)

* Save the imported/checked Stata panel for follow-up diagnostics.
save "`out'/wos_author_year_panel_sample.dta", replace

log close
