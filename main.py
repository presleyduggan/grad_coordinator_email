import pandas as pd
import os
import tkinter as tk
from tkinter.filedialog import asksaveasfilename


class NoCopyPaste:

    def __init__(self):
        self.replace_coord = {
        "Accounting - Advanced Certificate" : "accounting.gradcoord@unlv.edu",
        "Accounting - Certificate" : "accounting.gradcoord@unlv.edu",
        "Accounting M.S." : "accounting.gradcoord@unlv.edu",
        "Addiction Studies - Advanced Certificate" : "mentalhealth.gradcoord@unlv.edu",
        "Aerospace Engineering M.S.A.E." : "mechanical.gradcoord@unlv.edu",
        "Anthropology M.A." : "anthropology.gradcoord@unlv.edu",
        "Anthropology Ph.D." : "anthropology.gradcoord@unlv.edu",
        "Applied Economics and Data Intelligence M.S." : "daae.gradcoord@unlv.edu",
        "Architecture M.Arch." : "architecture.gradcoord@unlv.edu",
        "Artist Diploma A.D." : "music.gradcoord@unlv.edu",
        "Art M.F.A." : "art.gradcoord@unlv.edu",
        "Astronomy M.S." : "physics.gradcoord@unlv.edu",
        "Astronomy Ph.D." : "physics.gradcoord@unlv.edu",
        "Biobehavioral Nursing - Post Graduate Certificate" : "nursing.doc.director@unlv.edu",
        "Biochemistry M.S." : "chemistry.gradcoord@unlv.edu",
        "Biological Sciences M.S." : "lifesciences.gradcoord@unlv.edu",
        "Biological Sciences Ph.D" : "lifesciences.gradcoord@unlv.edu",
        "Biomedical Engineering M.S." : "mechanical.gradcoord@unlv.edu",
        "Business Administration - Certificate" : "mba.director@unlv.edu",
        "Business Administration - Executive M.B.A." : "emba.director@unlv.edu",
        "Business Administration M.B.A." : "mba.director@unlv.edu",
        "Business Administration M.B.A. / Civil & Environmental Engineering M.S.E. (Dual)" : "mba.director@unlv.edu",
        "Business Administration M.B.A. / Computer Science M.S.C.S. (Dual)" : "mba.director@unlv.edu",
        "Business Administration M.B.A. / Dentistry D.M.D. (Dual)" : "mba.director@unlv.edu",
        "Business Administration M.B.A. / Doctor of Medicine (Dual)" : "mba.director@unlv.edu",
        "Business Administration M.B.A. / Hotel Administration M.S. (Dual)" : "mba.director@unlv.edu",
        "Business Administration M.B.A. / Juris Doctorate (Dual)" : "mba.director@unlv.edu",
        "Business Administration M.B.A. / Management Information Systems M.S. (Dual)" : "met.gradcoord@unlv.edu",
        "Business Administration M.B.A. / Quantitative Finance M.S. (Dual)" : "mba.director@unlv.edu",
        "Career and Technical Education - Certificate" : "tl.gradcoord@unlv.edu",
        "Chemistry M.S." : "chemistry.gradcoord@unlv.edu",
        "Chemistry Ph.D." : "chemistry.gradcoord@unlv.edu",
        "Chief Diversity Officer in Higher Education - Certificate" : "highered.gradcoord@unlv.edu",
        "Civil & Environmental Engineering M.S.E." : "ceec.gradcoord@unlv.edu",
        "Civil & Environmental Engineering Ph.D." : "ceec.gradcoord@unlv.edu",
        "Clinical Mental Health Counseling M.S." : "mentalhealth.gradcoord@unlv.edu",
        "Clinical Psychology Ph.D." : "psychology.gradcoord@unlv.edu",
        "College Sport Leadership - Certificate" : "csl.cert.gradcoord@unlv.edu",
        "Communication Studies M.A." : "commstudies.gradcoord@unlv.edu",
        "Computer Science M.S.C.S." : "computerscience.gradcoord@unlv.edu",
        "Computer Science Ph.D." : "computerscience.gradcoord@unlv.edu",
        "Construction Management M.S." : "ceec.gradcoord@unlv.edu",
        "Couple & Family Therapy M.S." : "cft.director@unlv.edu",
        "Creative Writing M.F.A." : "creativewriting.gradcoord@unlv.edu",
        "Criminal Justice M.A." : "criminaljustice.gradcoord@unlv.edu",
        "Criminology and Criminal Justice Ph.D." : "criminaljustice.gradcoord@unlv.edu",
        "Crisis & Emergency Management M.S." : "ecm.gradcoord@unlv.edu",
        "Curriculum & Instruction Ed.D." : "tl.doc.gradcoord@unlv.edu",
        "Curriculum & Instruction Ed.S." : "tl.doc.gradcoord@unlv.edu", 
        "Curriculum & Instruction M.Ed." : "tl.gradcoord@unlv.edu",
        "Curriculum & Instruction M.S." : "tl.gradcoord@unlv.edu",
        "Curriculum & Instruction Ph.D." : "tl.doc.gradcoord@unlv.edu",
        "Cybersecurity M.S." : "cybersecurity.gradcoord@unlv.edu",
        "Cybersecurity M.S. / Management Information Systems M.S. (Dual)" : "cybersecurity.gradcoord@unlv.edu",
        "Data Analytics and Applied Economics M.S." : "daae.gradcoord@unlv.edu",
        "Data Analytics - Certificate" : "met.gradcoord@unlv.edu",
        "Data Analytics M.S." : "dataanalytics.gradcoord@unlv.edu",
        "Dentistry D.M.D." : "christine.ancajas@unlv.edu",
        "Design M.Des." : "mhid.gradcoord@unlv.edu",
        "Early Childhood Education M.Ed." : "ems.gradcoord@unlv.edu",
        "Early Childhood Special Education - Infancy - Certificate" : "ems.gradcoord@unlv.edu",
        "Early Childhood Special Education - Preschool - Certificate" : "ems.gradcoord@unlv.edu",
        "Economics M.A." : "economics.gradcoord@unlv.edu",
        "Economics M.A. / Mathematical Sciences M.S. (Dual)" : "economics.gradcoord@unlv.edu", 
        "Educational Policy and Leadership M.Ed." : "epl.gradcoord@unlv.edu",
        "Educational Psychology Ed.S." : "edpsych.gradcoord@unlv.edu",
        "Educational Psychology M.S." : "edpsych.gradcoord@unlv.edu",
        "Educational Psychology Ph.D." : "edpsych.gradcoord@unlv.edu",
        "Educational Psychology Ph.D. / Juris Doctorate (Dual)" : "edpsych.gradcoord@unlv.edu",
        "Electrical Engineering M.S." : "ecg.gradcoord@unlv.edu",
        "Electrical Engineering M.S. / Mathematical Sciences M.S. (Dual)" : "ecg.gradcoord@unlv.edu",
        "Electrical Engineering Ph.D." : "ecg.gradcoord@unlv.edu",
        "Electrical Engineering Ph.D. / Mathematical Sciences M.S. (Dual)" : "ecg.gradcoord@unlv.edu",
        "Elementary Teaching - Conditional Licensure Certificate" : "tl.gradcoord@unlv.edu",
        "Emergency & Crisis Management M.S." : "ecm.gradcoord@unlv.edu",
        "Emergency Crisis Management Cybersecurity - Certificate" : "ecm.gradcoord@unlv.edu",
        "Emergency Management Cybersecurity - Certificate" : "ecm.gradcoord@unlv.edu",
        "Emergency Nurse Practitioner - Advanced Certificate" : "nursing.masters.director@unlv.edu",
        "English Language Acquisition & Development (ELAD) - Certificate" : "ems.gradcoord@unlv.edu", 
        "English Language Learning M.Ed." : "ems.gradcoord@unlv.edu",
        "English M.A." : "english.gradcoord@unlv.edu",
        "English Ph.D." : "english.gradcoord@unlv.edu",
        "Environmental Science M.S." : "sppl.gradcoord@unlv.edu",
        "Environmental Science Ph.D." : "sppl.gradcoord@unlv.edu",
        "Executive Educational Leadership Ed.D." : "execedl.gradcoord@unlv.edu",
        "Exercise Physiology M.S." : "kinesiology.gradcoord@unlv.edu",
        "Family Nurse Practitioner - Advanced Certificate" : "nursing.masters.director@unlv.edu",
        "Finance - Certificate" : "finance.gradcoord@unlv.edu",
        "Gaming Management - Certificate" : "hospitalityadminmha.gradcoord@unlv.edu",
        "Geoscience M.S." : "geoscience.gradcoord@unlv.edu",
        "Geoscience Ph.D." : "geoscience.gradcoord@unlv.edu",
        "Global Teaching - Certificate" : "tl.gradcoord@unlv.edu",
        "Global Teaching Research - Certificate" : "tl.gradcoord@unlv.edu",
        "Healthcare Administration Executive M.H.A." : "healthcareadmin.gradcoord@unlv.edu",
        "Healthcare Administration M.H.A." : "healthcareadmin.gradcoord@unlv.edu",
        "Healthcare Interior Design M.H.I.D." : "mhid.gradcoord@unlv.edu",
        "Health Physics M.S." : "healthphysics.gradcoord@unlv.edu",
        "Higher Education - Certificate" : "highered.gradcoord@unlv.edu", 
        "Higher Education M.Ed." : "highered.gradcoord@unlv.edu",
        "Higher Education Ph.D." : "highered.gradcoord@unlv.edu",
        "Higher Education Ph.D. / Juris Doctorate (Dual)" : "highered.gradcoord@unlv.edu",
        "Hispanic Studies M.A." : "wlc.gradcoord@unlv.edu",
        "History M.A." : "history.gradcoord@unlv.edu",
        "History Ph.D." : "history.gradcoord@unlv.edu",
        "Hospitality Administration M.H.A." : "hospitalityadminmha.gradcoord@unlv.edu",
        "Hospitality Administration Ph.D." : "hospitalityadminphd.gradcoord@unlv.edu",
        "Hospitality Design - Certificate" : "architecture.gradcoord@unlv.edu",
        "Hotel Administration M.S." : "hoteladminms.gradcoord@unlv.edu",
        "Hotel Administration M.S. / Management Information Systems M.S. (Dual)" : "hoteladminmsdual.gradcoord@unlv.edu",
        "Infection Prevention - Certificate" : "eoh.gradcoord@unlv.edu", 
        "Intercollegiate and Professional Sport Management M.Ed." : "highered.gradcoord@unlv.edu",
        "Interdisciplinary Health Sciences Ph.D." : "ihs.gradcoord@unlv.edu",
        "Journalism & Media Studies M.A." : "jms.gradcoord@unlv.edu",
        "K-8 Integrated STEM Education - Certificate" : "tl.gradcoord@unlv.edu",
        "Kinesiology M.S." : "kinesiology.gradcoord@unlv.edu",
        "Kinesiology Ph.D." : "kinesiology.gradcoord@unlv.edu",
        "Leadership for Teachers and Professionals - Certificate" : "tl.gradcoord@unlv.edu",
        "Leadership in English Language Acquisition & Development (LELAD) - Certificate" : "ems.gradcoord@unlv.edu",
        "Learning and Technology Ph.D." : "edpsych.gradcoord@unlv.edu",
        "Learning Sciences Ph.D." : "edpsych.gradcoord@unlv.edu",
        "Management - Certificate" : "mgt.cert.gradcoord@unlv.edu",
        "Management Information Systems - Certificate" : "met.gradcoord@unlv.edu",
        "Management Information Systems M.S." : "met.gradcoord@unlv.edu",
        "Marriage & Family Therapy M.S." : "cft.director@unlv.edu",
        "Materials & Nuclear Engineering M.S." : "mechanical.gradcoord@unlv.edu",
        "Mathematical Sciences M.S." : "math.gradcoord@unlv.edu",
        "Mathematical Sciences Ph.D." : "math.gradcoord@unlv.edu",
        "Mechanical Engineering M.S." : "mechanical.gradcoord@unlv.edu",
        "Mechanical Engineering Ph.D." : "mechanical.gradcoord@unlv.edu",
        "Medical Physics - Advanced Certificate" : "healthphysics.gradcoord@unlv.edu",
        "Medical Physics D.M.P." : "healthphysics.gradcoord@unlv.edu",
        "Mental Health Counseling - Advanced Certificate" : "mentalhealth.gradcoord@unlv.edu",
        "Multicultural Education - Certificate" : "mce.cert.gradcoord@unlv.edu",
        "Music M.M." : "music.gradcoord@unlv.edu",
        "Neuroscience Ph.D." : "neuroscience.director@unlv.edu",
        "New Venture Management - Certificate" : "nvm.cert.gradcoord@unlv.edu",
        "Non-Profit Management - Certificate" : "sppl.gradcoord@unlv.edu",
        "Nuclear Criticality Safety Engineering - Certificate" : "mechanical.gradcoord@unlv.edu",
        "Nuclear Safeguards and Security - Certificate" : "mechanical.gradcoord@unlv.edu",
        "Nursing D.N.P." : "nursing.dnp.director@unlv.edu",
        "Nursing Education - Advanced Certificate" : "nursing.masters.director@unlv.edu",
        "Nursing Education Exchange NEXus" : "gradcollege@unlv.edu",
        "Nursing M.S." : "nursing.masters.director@unlv.edu",
        "Nursing Ph.D." : "nursing.doc.director@unlv.edu",
        "Nutrition Sciences M.S." : "nutritionsci.gradcoord@unlv.edu",
        "Occupational Therapy O.T.D." : "otd.gradcoord@unlv.edu",
        "Online Teaching and Training - Certificate" : "tl.gradcoord@unlv.edu",
        "Oral Biology M.S." : "oralbiology.gradcoord@unlv.edu",
        "Oral Biology Ph.D." : "oralbiology.gradcoord@unlv.edu",
        "Pediatric Nurse Practitioner - Advanced Certificate" : "nursing.masters.director@unlv.edu",
        "Performance D.M.A." : "music.gradcoord@unlv.edu",
        "Physical Therapy D.P.T." : "dpt.gradcoord@unlv.edu",
        "Physics M.S." : "physics.gradcoord@unlv.edu",
        "Physics Ph.D." : "physics.gradcoord@unlv.edu",
        "Political Science M.A." : "psc.gradcoord@unlv.edu",
        "Political Science Ph.D." : "psc.gradcoord@unlv.edu",
        "Post-Professional Occupational Therapy O.T.D." : "otd.gradcoord@unlv.edu",
        "Program Evaluation and Assessment - Certificate" : "edpsych.gradcoord@unlv.edu",
        "Psychiatric Mental Health Nurse Practitioner - Advanced Certificate" : "nursing.cert.coordinator@unlv.edu",
        "Psychiatric Mental Health Nurse Practitioner for FNP - Advanced Certificate" : "nursing.cert.coordinator@unlv.edu",
        "Psychological and Brain Sciences Ph.D." : "psychbrainsc.gradcoord@unlv.edu",
        "Psychology M.A." : "psychology.gradcoord@unlv.edu",
        "Psychology Ph.D." : "psychology.gradcoord@unlv.edu",
        "Public Administration M.P.A." : "sppl.gradcoord@unlv.edu",
        "Public Affairs Ph.D." : "sppl.gradcoord@unlv.edu",
        "Public Health - Certificate" : "eoh.gradcoord@unlv.edu",
        "Public Health M.P.H." : "eoh.gradcoord@unlv.edu",
        "Public Health M.P.H. / Dentistry D.M.D. (Dual)" : "eoh.gradcoord@unlv.edu",
        "Public Health Ph.D." : "eoh.gradcoord@unlv.edu",
        "Public Management - Certificate" : "sppl.gradcoord@unlv.edu",
        "Public Policy D.P.P." : "dpp.gradcoord@unlv.edu",
        "Quantitative Finance M.S." : "fin.director@unlv.edu",
        "Quantitative Psychology - Certificate" : "quantpsy.cert.gradcoord@unlv.edu",
        "Radiochemistry Ph.D." : "radchem.gradcoord@unlv.edu",
        "Regional Professional Development" : "Program RPDP	gradcollege@unlv.edu",
        "School Counseling M.Ed." : "schoolcounseling.gradcoord@unlv.edu",
        "School Psychology Ed.S." : "schoolpsych.gradcoord@unlv.edu",
        "School Psychology Ph.D." : "schoolpsych.gradcoord@unlv.edu",
        "Secondary Teaching - Conditional Licensure Certificate" : "tl.gradcoord@unlv.edu",
        "Social Justice Studies - Certificate" : "tl.gradcoord@unlv.edu",
        "Social Science Methods - Certificate" : "ssm.cert.gradcoord@unlv.edu",
        "Social Work M.S.W." : "socialwork.gradcoord@unlv.edu",
        "Social Work M.S.W. / Juris Doctorate (Dual)" : "socialwork.gradcoord@unlv.edu",
        "Sociology M.A." : "sociology.gradcoord@unlv.edu",
        "Sociology Ph.D." : "sociology.gradcoord@unlv.edu",
        "Solar and Renewable Energy - Certificate" : "solar.cert.gradcoord@unlv.edu",
        "Spanish Translation - Certificate" : "wlc.gradcoord@unlv.edu",
        "Special Education - Certificate" : "ems.gradcoord@unlv.edu", 
        "Special Education M.Ed." : "ems.gradcoord@unlv.edu",
        "Special Education Ph.D." : "ems.doc.gradcoord@unlv.edu",
        "Teacher Education Ph.D." : "teachered.gradcoord@unlv.edu",
        "Special Education Ph.D. / Juris Doctorate (Dual)" : "ems.doc.gradcoord@unlv.edu",
        "Teacher Licensure: K-12 Music - Certificate" : "music.gradcoord@unlv.edu",
        "Theatre Arts M.A." : "theatre.gradcoord@unlv.edu",
        "Theatre Arts M.F.A." : "theatre.gradcoord@unlv.edu",
        "Transportation M.S.T." : "ceec.gradcoord@unlv.edu",
        "Urban Leadership M.A." : "sppl.gradcoord@unlv.edu",
        "Water Resources Management M.S." : "wrm.director@unlv.edu",
        "Workforce Development and Organizational Leadership Ph.D." : "sppl.gradcoord@unlv.edu",
        "Writing for Dramatic Media M.F.A." : "film.gradcoord@unlv.edu",
        "Writing for Dramatic Media - Certificate" : "film.gradcoord@unlv.edu"
        }

        self.window = tk.Tk()

        self.window.title("Don't Copy Paste Emails")

        self.window.geometry('600x600')

        lbl = tk.Label(self.window, text="Name row to be replaced 'CoordEmail' and have academic plan names in it")

        lbl.place(x=75, y=100)

        btn = tk.Button(self.window, text="Feed Me Excel Spreadsheet", command=self.read_data, fg='blue')

        btn.place(x=200, y=400)

        self.window.mainloop()

    def read_data(self):
        filename = tk.filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
        self.data = pd.read_excel(filename)
        self.replace_data()
       

    def replace_data(self):
        self.data = self.data.replace({"CoordEmail": self.replace_coord})
        self.save_file()

    def save_file(self):
        filename = asksaveasfilename(
                defaultextension='.xlsx', filetypes=[("excel files","*.xlsx")],
                title="Choose filename")
        if not filename:
            return
        print(filename)
        self.data.to_excel(filename)



def main():
    test = NoCopyPaste()

if __name__ == "__main__":
    main()