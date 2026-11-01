from pathlib import Path
p=Path('build_model.py');t=p.read_text();t=t.replace("('Market_Rooms',96,","('Market_Rooms',108,").replace("('Social_Rooms',24,","('Social_Rooms',12,").replace("('Required_Social_Quota',.2,","('Required_Social_Quota',.1,")
t=t.replace('Selected municipality Ixelles','Selected municipality City of Brussels').replace('20% conservative management policy','Preliminary 10% policy; iteration 2 tests published City guidance')
t=t.replace("for name in ['Planning", "legal.append(['S8 Current guidance publication','City website lists the 2021 student housing recommendations. Municipal guidance, not a universal regulation.','https://www.bruxelles.be/plans-reglements-urbanisme','Current publication listing','2026-10-29','City of Brussels selected; exact permit conditions still unverified.'])\nfor name in ['Planning")
p.write_text(t)
p=Path('validate_finish.py');t=p.read_text().replace('selected municipality Ixelles','selected municipality City of Brussels');p.write_text(t)
