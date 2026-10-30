/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { MoveLeft, Download, FileText, CheckCircle2, RotateCw } from 'lucide-react';

// Declaration for docx library which is loaded via CDN
declare global {
  interface Window {
    docx: any;
    saveAs: any;
  }
}

interface FormData {
  role: string;
  ideesRecues: string;
  styleForces: string;
  axe1: string;
  axe2: string;
  visionInterface: string;
}

const INITIAL_DATA: FormData = {
  role: '',
  ideesRecues: '',
  styleForces: '',
  axe1: '',
  axe2: '',
  visionInterface: '',
};

export default function App() {
  const [formData, setFormData] = useState<FormData>(INITIAL_DATA);
  const [saveStatus, setSaveStatus] = useState<'saved' | 'saving'>('saved');
  const [isLoaded, setIsLoaded] = useState(false);
  const saveTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  // Load from localStorage on mount
  useEffect(() => {
    const saved = localStorage.getItem('nexia_identity_sheet');
    if (saved) {
      try {
        setFormData(JSON.parse(saved));
      } catch (e) {
        console.error('Error loading saved data:', e);
      }
    }
    setIsLoaded(true);
  }, []);

  // Auto-save logic
  const handleInputChange = (field: keyof FormData, value: string) => {
    const newData = { ...formData, [field]: value };
    setFormData(newData);
    setSaveStatus('saving');

    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);

    saveTimeoutRef.current = setTimeout(() => {
      localStorage.setItem('nexia_identity_sheet', JSON.stringify(newData));
      setSaveStatus('saved');
    }, 1000);
  };

  // Progress calculation
  const totalFields = Object.keys(INITIAL_DATA).length;
  const completedFields = (Object.values(formData) as string[]).filter(v => v.trim().length > 0).length;
  const progressPercent = Math.round((completedFields / totalFields) * 100);

  // Export to Word
  const exportToWord = async () => {
    // Check for global docx object
    const docxLib = (window as any).docx;
    const saveAsLib = (window as any).saveAs;

    if (!docxLib || !saveAsLib) {
      alert("Le chargement des outils d'exportation est en cours. Veuillez patienter quelques secondes.");
      return;
    }

    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docxLib;

    try {
      const doc = new Document({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({
                text: "Fiche d'identité managériale — NEXIA",
                heading: HeadingLevel.TITLE,
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
              }),
              new Paragraph({
                text: `Module 1 — ${new Date().toLocaleDateString('fr-FR')}`,
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 },
              }),

              ...createRubriqueWord(1, "Mon rôle chez NEXIA", formData.role || "Non rempli"),
              ...createRubriqueWord(2, "Les idées reçues que je dois surveiller", formData.ideesRecues || "Non rempli"),
              ...createRubriqueWord(3, "Mon style naturel et mes forces", formData.styleForces || "Non rempli"),
              ...createRubriqueWord(4, "Mes 2 axes de développement prioritaires", `Axe 1: ${formData.axe1 || "Non rempli"}\n\nAxe 2: ${formData.axe2 || "Non rempli"}`),
              ...createRubriqueWord(5, "Ma vision du rôle d'interface", formData.visionInterface || "Non rempli"),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      saveAsLib(blob, `Fiche_Identite_Manageriale_NEXIA_${new Date().toISOString().split('T')[0]}.docx`);
    } catch (error) {
      console.error("Erreur lors de l'export Word:", error);
      alert("Une erreur est survenue lors de la génération du fichier Word.");
    }
  };

  const createRubriqueWord = (num: number, title: string, content: string) => {
    const docxLib = (window as any).docx;
    const { Paragraph, TextRun, HeadingLevel } = docxLib;
    
    return [
      new Paragraph({
        text: `Rubrique ${num} — ${title}`,
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 400, after: 200 },
      }),
      new Paragraph({
        children: content.split('\n').map((line, i, arr) => (
          new TextRun({
            text: line,
            break: i < arr.length - 1 ? 1 : 0,
          })
        )),
        spacing: { after: 200 },
      }),
    ];
  };

  const exportToPDF = () => {
    window.print();
  };

  if (!isLoaded) return null;

  return (
    <div className="min-h-screen flex flex-col bg-[#fdfdfd] font-sans text-[var(--color-text-primary)]">
      <style>{`
        :root {
          --color-accent: #d41875;
          --color-accent-light: #fbeaf4;
          --color-accent-dark: #9e1056;
          --color-text-primary: #2d2d2d;
          --color-text-secondary: #666666;
          --color-background-secondary: #f9f9f9;
          --color-border-tertiary: #e0e0e0;
          --border-radius-md: 8px;
          --border-radius-lg: 12px;
          --font-sans: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        }

        @media screen {
          .print-only { display: none; }
        }

        @media print {
          .no-print { display: none !important; }
          body { background: white !important; }
          .min-h-screen { height: auto !important; min-height: 0 !important; }
          .card { border: none !important; margin-bottom: 2rem !important; padding: 0 !important; }
          .card h2 { font-size: 1.25rem !important; color: black !important; border-bottom: 2px solid #eee !important; padding-bottom: 0.5rem !important; margin-bottom: 1rem !important; }
          .card p.consigne { display: none !important; }
          .textarea-print { white-space: pre-wrap; display: block !important; border: 1px solid #eee; padding: 1rem; border-radius: 4px; }
          textarea { display: none !important; }
          .rubrique-reference { display: none !important; }
        }

        textarea {
          transition: border-color 0.2s, box-shadow 0.2s;
        }

        textarea:focus {
          outline: none;
          border-color: var(--color-accent);
          box-shadow: 0 0 0 2px var(--color-accent-light);
        }

        .card {
          background: white;
          border: 0.5px solid var(--color-border-tertiary);
          border-radius: var(--border-radius-md);
          padding: 18px;
          display: flex;
          flex-direction: column;
          gap: 12px;
        }

        .sr-only {
          position: absolute;
          width: 1px;
          height: 1px;
          padding: 0;
          margin: -1px;
          overflow: hidden;
          clip: rect(0, 0, 0, 0);
          border: 0;
        }

        button {
          transition: opacity 0.2s, transform 0.1s;
          cursor: pointer;
        }

        button:active {
          transform: scale(0.98);
        }
      `}</style>

      {/* SR-ONLY Navigation Landmark */}
      <h2 className="sr-only">Fiche d'identité managériale — Module 1 NEXIA. Remplissez les 5 rubriques et exportez votre fiche.</h2>

      {/* Header */}
      <header className="no-print sticky top-0 z-20 bg-white border-b border-[var(--color-border-tertiary)] px-8 h-20 flex items-center justify-between">
        <div className="flex items-center gap-5">
          <a 
            href="https://1001parcours.cci.fr/experience/mod/page/view.php?id=2258" 
            className="flex items-center gap-1.5 text-sm text-[var(--color-text-secondary)] hover:text-[var(--color-accent)] transition-colors text-[14px]"
            aria-label="Retour au parcours"
          >
            <span>←</span>
            <span>Retour au parcours</span>
          </a>
          <div className="bg-[var(--color-accent-light)] text-[var(--color-accent)] text-[12px] font-medium px-2.5 py-1 rounded-[4px] uppercase tracking-wider">Module 1</div>
          <div className="flex flex-col">
            <h1 className="text-[18px] font-medium leading-tight">Fiche d'identité managériale</h1>
            <span className="text-[13px] text-[var(--color-text-secondary)]">Livrable M1 — NEXIA</span>
          </div>
        </div>

        <div 
          className="flex items-center gap-2 text-[13px] text-[var(--color-text-secondary)]" 
          aria-live="polite"
        >
          <AnimatePresence mode="wait">
            {saveStatus === 'saved' ? (
              <motion.div 
                key="saved"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="flex items-center gap-2"
              >
                <div className="w-2 h-2 rounded-full bg-[#4caf50]"></div>
                <span>Sauvegardé</span>
              </motion.div>
            ) : (
              <motion.div 
                key="saving"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="flex items-center gap-2"
              >
                <div className="w-2 h-2 rounded-full bg-orange-400 animate-pulse"></div>
                <span>Sauvegarde…</span>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </header>

      {/* Main Content - Minimalist Grid */}
      <main className="flex-1 p-8 grid grid-cols-1 lg:grid-cols-2 gap-6 pb-24 overflow-y-auto">
        {/* Left Column */}
        <div className="flex flex-col gap-6">
          {/* Rubrique 1 */}
          <section id="r1" aria-labelledby="t1" className="card">
            <h2 id="t1" className="text-[15px] font-medium text-[var(--color-accent)]">1. Mon rôle chez NEXIA</h2>
            <p className="consigne text-[12.5px] text-[var(--color-text-secondary)] leading-normal">
              Décrivez votre périmètre de responsabilité dans la situation NEXIA. Précisez vos trois missions principales dans ce contexte.
            </p>
            <div className="relative flex-1">
              <label htmlFor="role-input" className="sr-only">Description de votre rôle</label>
              <textarea
                id="role-input"
                className="w-full h-32 p-3 text-[13px] border border-[var(--color-border-tertiary)] rounded-md resize-none"
                placeholder=""
                value={formData.role}
                onChange={(e) => handleInputChange('role', e.target.value)}
              />
              <div className="print-only textarea-print">{formData.role || "Non rempli"}</div>
            </div>
          </section>

          {/* Rubrique 2 */}
          <section id="r2" aria-labelledby="t2" className="card">
            <h2 id="t2" className="text-[15px] font-medium text-[var(--color-accent)]">2. Les idées reçues à surveiller</h2>
            <p className="consigne text-[12.5px] text-[var(--color-text-secondary)] leading-normal">
              Identifiez les 2 ou 3 idées reçues qui pourraient influencer votre propre posture. Expliquez comment elles pourraient se manifester dans votre management.
            </p>
            
            <div className="rubrique-reference bg-[var(--color-background-secondary)] rounded-md p-3 border border-[var(--color-border-tertiary)]">
              <ol className="text-[11px] text-[var(--color-text-secondary)] list-decimal list-inside space-y-1">
                <li>Un bon manager est d'abord un bon expert</li>
                <li>Manager c'est décider seul</li>
                <li>Un bon manager est apprécié de tous</li>
                <li>Le management ça s'improvise</li>
                <li>Manager c'est contrôler</li>
              </ol>
            </div>

            <div className="relative flex-1">
              <label htmlFor="idees-input" className="sr-only">Vos réflexions sur les idées reçues</label>
              <textarea
                id="idees-input"
                className="w-full h-40 p-3 text-[13px] border border-[var(--color-border-tertiary)] rounded-md resize-none"
                placeholder=""
                value={formData.ideesRecues}
                onChange={(e) => handleInputChange('ideesRecues', e.target.value)}
              />
              <div className="print-only textarea-print">{formData.ideesRecues || "Non rempli"}</div>
            </div>
          </section>

          {/* Rubrique 3 (Adapted to column 1 to balance grid) */}
          <section id="r3" aria-labelledby="t3" className="card">
            <h2 id="t3" className="text-[15px] font-medium text-[var(--color-accent)]">3. Mon style naturel et mes forces</h2>
            <p className="consigne text-[12.5px] text-[var(--color-text-secondary)] leading-normal">
              Décrivez votre style dominant et ses forces. Mettez-les en lien avec l'équipe NEXIA et justifiez vos réponses. Pour quels collaborateurs votre style est-il un atout ? Pour lesquels devrez-vous l'ajuster ?
            </p>
            <div className="relative flex-1">
              <label htmlFor="style-input" className="sr-only">Diagnostic de style et forces</label>
              <textarea
                id="style-input"
                className="w-full h-32 p-3 text-[13px] border border-[var(--color-border-tertiary)] rounded-md resize-none"
                placeholder=""
                value={formData.styleForces}
                onChange={(e) => handleInputChange('styleForces', e.target.value)}
              />
              <div className="print-only textarea-print">{formData.styleForces || "Non rempli"}</div>
            </div>
          </section>
        </div>

        {/* Right Column */}
        <div className="flex flex-col gap-6">
          {/* Rubrique 4 */}
          <section id="r4" aria-labelledby="t4" className="card">
            <h2 id="t4" className="text-[15px] font-medium text-[var(--color-accent)]">4. Mes 2 axes de développement</h2>
            <p className="consigne text-[12.5px] text-[var(--color-text-secondary)] leading-normal">
              Formulez 2 axes concrets, en lien avec les défis de l'équipe NEXIA.
            </p>
            
            <div className="space-y-4 pt-2">
              <div className="relative">
                <label htmlFor="axe1-input" className="block text-[12px] font-medium text-[var(--color-accent-dark)] mb-1">Axe 1</label>
                <textarea
                  id="axe1-input"
                  className="w-full h-24 p-3 text-[13px] border border-[var(--color-border-tertiary)] rounded-md resize-none"
                  placeholder="Ex : Je dois développer ma capacité à [action]..."
                  value={formData.axe1}
                  onChange={(e) => handleInputChange('axe1', e.target.value)}
                  maxLength={500}
                />
                <div className="no-print text-[10px] text-gray-400 text-right mt-1">
                  {formData.axe1.length} / 500
                </div>
              </div>

              <div className="relative">
                <label htmlFor="axe2-input" className="block text-[12px] font-medium text-[var(--color-accent-dark)] mb-1">Axe 2</label>
                <textarea
                  id="axe2-input"
                  className="w-full h-24 p-3 text-[13px] border border-[var(--color-border-tertiary)] rounded-md resize-none"
                  placeholder="Ex : Je dois développer ma capacité à [action]..."
                  value={formData.axe2}
                  onChange={(e) => handleInputChange('axe2', e.target.value)}
                  maxLength={500}
                />
                <div className="no-print text-[10px] text-gray-400 text-right mt-1">
                  {formData.axe2.length} / 500
                </div>
              </div>
            </div>
          </section>

          {/* Rubrique 5 */}
          <section id="r5" aria-labelledby="t5" className="card flex-1">
            <h2 id="t5" className="text-[15px] font-medium text-[var(--color-accent)]">5. Ma vision du rôle d'interface</h2>
            <p className="consigne text-[12.5px] text-[var(--color-text-secondary)] leading-normal">
              Comment concevez-vous votre rôle entre l'équipe NEXIA et la direction ? Qu'est-ce que vous vous engagez à faire remonter ? À quelle fréquence ? Avec quelle posture ?
            </p>
            <div className="relative flex-1">
              <label htmlFor="vision-input" className="sr-only">Votre conception du rôle d'interface</label>
              <textarea
                id="vision-input"
                className="w-full h-48 p-3 text-[13px] border border-[var(--color-border-tertiary)] rounded-md resize-none"
                placeholder=""
                value={formData.visionInterface}
                onChange={(e) => handleInputChange('visionInterface', e.target.value)}
              />
              <div className="print-only textarea-print">{formData.visionInterface || "Non rempli"}</div>
            </div>
          </section>
        </div>
      </main>

      {/* Export Zone */}
      <footer className="no-print h-[72px] bg-[var(--color-background-secondary)] border-t border-[var(--color-border-tertiary)] px-8 flex items-center justify-end gap-4 z-20">
        <p className="flex-1 text-[14px] font-medium text-[var(--color-text-primary)]">Exporter ma fiche :</p>
        <button
          onClick={exportToPDF}
          className="px-6 py-2.5 rounded-md border border-[var(--color-border-tertiary)] bg-white text-[14px] font-medium text-[var(--color-text-primary)] hover:opacity-80"
        >
          Télécharger en PDF
        </button>
        <button
          onClick={exportToWord}
          className="px-6 py-2.5 rounded-md bg-[var(--color-accent)] text-white text-[14px] font-medium hover:opacity-80"
        >
          Télécharger en Word
        </button>
      </footer>
    </div>
  );
}
