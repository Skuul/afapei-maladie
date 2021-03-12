from .commun import *
from .commun import Enregistrement, Salarie

Absence = namedtuple('Absence', ['debut', 'fin', 'motif', 'heures', 'heures_3_premiers_jours', 'code_service'])

def formatAbsence(absence:Enregistrement):
    if absence.Motif == 'MAL':
        return Absence(absence.Date_planning, absence.Date_planning, absence.Motif, absence.Heures, absence.Heures, absence.Code_service)
    else:
        return Absence(absence.Date_planning, absence.Date_planning, absence.Motif, absence.Heures, '', absence.Code_service)

def f(absencesSalarie):
    l = []
    for xs in absencesSalarie:
        l.extend(xs)
    return l

def _getAbsencesAvecServices(absences: List[Enregistrement]) -> Dict[str, Absence]:
    absencesSalarie = defaultdict(list)
    for absence in absences:
        absencesService = absencesSalarie[absence.Code_service]
        if not len(absencesService):
            absencesService.append(formatAbsence(absence))
            continue
        else:
            derniereAbsence = absencesService.pop()
            estMemeMotif = derniereAbsence.motif == absence.Motif
            if estMemeMotif and estJourSuivant(derniereAbsence.fin, absence.Date_planning):
                derniereAbsence = derniereAbsence._replace(fin=absence.Date_planning, heures = (derniereAbsence.heures + absence.Heures))

                if absence.Motif == 'MAL' and pasPlusDeTroisJours(derniereAbsence.debut, absence.Date_planning):
                    derniereAbsence = derniereAbsence._replace(heures_3_premiers_jours=(derniereAbsence.heures_3_premiers_jours + absence.Heures))

                absencesService.append(derniereAbsence)
            else:
                absencesService.append(derniereAbsence)

                # Ajouter une absence
                absencesService.append(formatAbsence(absence))

    return f(dict(absencesSalarie).values())

def _outputXlsx(absencesSalaries:Dict[Salarie, Absence], filename):
    wb = openpyxl.Workbook()
    wb.iso_dates = True
    ws = wb.active
    ws.title = 'Absences par salarié'
    
    # Headers
    ws.append((
        'Matricule',
        'Nom',
        'Prénom',
        'Service',
        'Début',
        'Fin',
        'Motif',
        'Heures',
        'Nombre de jours',
        'Heures pendant les 3 premiers jours', # MAL uniquement
        'Nombre de jours comptés' # MAL uniquement
    ))

    # Data
    for salarie, absences in absencesSalaries.items():
        for absence in absences:
            nbJours = (absence.fin - absence.debut).days+1

            # Absence maladie, infos supplémentaires
            if absence.motif == 'MAL':
                ws.append((
                    salarie.matricule,          # 0
                    salarie.nom,# nom
                    salarie.prenom,# prénom
                    absence.code_service,
                    *(absence[:4]),     # 1:4
                    nbJours,            # 5
                    absence[4],         # 6
                    min(nbJours, 3)     #7
                ))
            else:
                ws.append((
                    salarie.matricule,          # 0
                    salarie.nom,# nom
                    salarie.prenom,# prénom
                    absence.code_service,
                    *(absence[:4]),   # 1:4
                    nbJours,    # 5
                    None,       # 6
                    None        # 7
                ))
    
    wb.save(filename)

def algo(filename:str, outDir):
    wb = openpyxl.load_workbook(filename)
    sheetName = wb.sheetnames[0]
    sheet = wb[sheetName]
    
    base = []

    it = iter(sheet.rows)
    # Ignore header
    next(it)

    for row in it:
        base.append(Enregistrement(*(cell.value for cell in row)))

    absencesSalaries = _processBase(base)
    outputFileName = fs_sep.join([outDir, filename.split(fs_sep)[-1].split(".")[0]]) + '.xlsx'
    _outputXlsx(absencesSalaries, outputFileName)
    print(f'Done with {filename.split(fs_sep)[-1]}')
    # Write to file

def _processBase(base: List[Enregistrement]):
    absencesSalaries = dict()

    for salarie, absences in filterAbsences(base):
        print(f'Looking at {salarie}...')
        absencesSalaries[salarie] = _getAbsencesAvecServices(absences)

    return absencesSalaries