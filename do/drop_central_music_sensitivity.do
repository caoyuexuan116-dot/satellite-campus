version 17.0
clear all
set more off

local root "C:/Users/Dell/Documents/学习wos数据使用"
local out "`root'/outputs"

log using "`out'/drop_central_music_sensitivity.log", replace text

use "`out'/wos_author_year_panel_sample.dta", clear

di as text "Dropping Central Conservatory of Music from the strict-control sample"
drop if school_cn == "中央音乐学院"

di as text "Sample checks after dropping Central Conservatory of Music"
tab school_cn treated, missing
tab year treated, missing
tab treated post, missing
assert inrange(year, 2000, 2022)
assert treated == 0 | treated == 1
assert post == 0 if treated == 0
assert did == treated * post

preserve
keep school_cn school_id person_id
duplicates drop
bysort school_cn: gen n_person = _N
bysort school_cn: keep if _n == 1
list school_cn n_person, clean
restore

estimates clear

xtreg pub_count did i.year, fe vce(cluster school_id)
estimates store m_pub

xtreg cites_wos_core_sum did i.year, fe vce(cluster school_id)
estimates store m_cite_core

xtreg cites_all_db_sum did i.year, fe vce(cluster school_id)
estimates store m_cite_all

xtreg pub_count did i.year if incumbent_treated_author == 1, fe vce(cluster school_id)
estimates store m_pub_incumbent

estimates table m_pub m_cite_core m_cite_all m_pub_incumbent, b(%9.4f) se(%9.4f) stats(N r2_w)

log close
