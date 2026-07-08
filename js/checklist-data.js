/**
 * checklist-data.js
 * Bilingual EN/ES checklist content model for all 5 inspection phases.
 * Loaded as a plain script (not fetch/JSON) so the app works when opened
 * directly via file:// on iOS Safari without CORS/fetch restrictions.
 *
 * A standalone copy of this same content also ships as data/checklist-data.json
 * for use outside the app (e.g. importing into another system).
 */

const CHECKLIST_DATA = {
  "project": {
    "name_en": "Verizon North Royalton UPS Project",
    "name_es": "Proyecto UPS Verizon North Royalton",
    "subcontractor_en": "Americano Industrial Services",
    "subcontractor_es": "Americano Industrial Services",
    "general_contractor_en": "Key Construction",
    "general_contractor_es": "Key Construction",
    "scope_en": "General Construction / Rough Carpentry / Finishes support work",
    "scope_es": "Trabajo de apoyo en Construcción General / Carpintería en Bruto / Acabados"
  },
  "disclaimer_en": "This checklist is a project coordination and field inspection tool only. It does not modify, expand, or reduce the contractual scope of work. All work remains subject to the executed subcontract, contract documents, approved drawings, specifications, and direction from Key Construction.",
  "disclaimer_es": "Esta lista de verificación es únicamente una herramienta de coordinación de proyecto e inspección de campo. No modifica, amplía ni reduce el alcance contractual del trabajo. Todo el trabajo permanece sujeto al subcontrato ejecutado, los documentos contractuales, los planos aprobados, las especificaciones y las instrucciones de Key Construction.",
  "statusOptions": [
    {
      "id": "pass",
      "label_en": "Pass",
      "label_es": "Aprobado"
    },
    {
      "id": "fail",
      "label_en": "Fail",
      "label_es": "Rechazado"
    },
    {
      "id": "pending",
      "label_en": "Pending",
      "label_es": "Pendiente"
    },
    {
      "id": "na",
      "label_en": "Not Applicable",
      "label_es": "No Aplica"
    }
  ],
  "phases": [
    {
      "id": "phase1",
      "number": 1,
      "name_en": "Site Preparation and Work Area Protection",
      "name_es": "Preparación del Sitio y Protección del Área de Trabajo",
      "items": [
        {
          "id": "p1_i1",
          "en": "Build platforms over batteries and panels.",
          "es": "Construir plataformas sobre baterías y paneles."
        },
        {
          "id": "p1_i2",
          "en": "Provide static-free protection over batteries.",
          "es": "Proporcionar protección antiestática sobre las baterías."
        },
        {
          "id": "p1_i3",
          "en": "Install temporary separation barriers.",
          "es": "Instalar barreras de separación temporales."
        },
        {
          "id": "p1_i4",
          "en": "Provide work area separation.",
          "es": "Proporcionar separación del área de trabajo."
        },
        {
          "id": "p1_i5",
          "en": "Protect existing batteries, panels, flooring, ceiling grid, and adjacent Verizon equipment.",
          "es": "Proteger las baterías, paneles, pisos, rejilla de cielo raso y equipo de Verizon adyacente existentes."
        },
        {
          "id": "p1_i6",
          "en": "Confirm area is ready to release to the electrical contractor.",
          "es": "Confirmar que el área está lista para ser entregada al contratista eléctrico."
        }
      ]
    },
    {
      "id": "phase2",
      "number": 2,
      "name_en": "Firestopping",
      "name_es": "Sellado Contra Incendios",
      "items": [
        {
          "id": "p2_i1",
          "en": "Fire caulking new conduit penetrations.",
          "es": "Sellado contra incendios (calafateo) de las penetraciones nuevas de conduit."
        },
        {
          "id": "p2_i2",
          "en": "Verify fire-rated assemblies are maintained.",
          "es": "Verificar que se mantengan los conjuntos con clasificación contra incendios."
        },
        {
          "id": "p2_i3",
          "en": "Confirm all conduit penetrations are sealed.",
          "es": "Confirmar que todas las penetraciones de conduit estén selladas."
        },
        {
          "id": "p2_i4",
          "en": "Confirm sealant/material used is appropriate for the application.",
          "es": "Confirmar que el sellador/material utilizado sea el adecuado para la aplicación."
        },
        {
          "id": "p2_i5",
          "en": "Document each completed firestop location with notes and photos.",
          "es": "Documentar cada ubicación de sellado contra incendios terminada con notas y fotografías."
        }
      ]
    },
    {
      "id": "phase3",
      "number": 3,
      "name_en": "Architectural Restoration",
      "name_es": "Restauración Arquitectónica",
      "items": [
        {
          "id": "p3_i1",
          "en": "Patch and paint existing pipe penetrations.",
          "es": "Reparar y pintar las penetraciones de tuberías existentes."
        },
        {
          "id": "p3_i2",
          "en": "Patch and paint drywall as required.",
          "es": "Reparar y pintar el tablaroca (drywall) según sea necesario."
        },
        {
          "id": "p3_i3",
          "en": "Replace ceiling tiles to match existing conditions.",
          "es": "Reemplazar las losetas de cielo raso para igualar las condiciones existentes."
        },
        {
          "id": "p3_i4",
          "en": "Confirm finish, color, texture, and appearance match existing adjacent areas.",
          "es": "Confirmar que el acabado, color, textura y apariencia coincidan con las áreas adyacentes existentes."
        },
        {
          "id": "p3_i5",
          "en": "Verify no construction damage remains visible.",
          "es": "Verificar que no quede daño de construcción visible."
        }
      ]
    },
    {
      "id": "phase4",
      "number": 4,
      "name_en": "Floor Restoration",
      "name_es": "Restauración de Pisos",
      "items": [
        {
          "id": "p4_i1",
          "en": "Patch anchor holes in floors.",
          "es": "Reparar los orificios de anclaje en los pisos."
        },
        {
          "id": "p4_i2",
          "en": "Touch up epoxy flooring in the work area.",
          "es": "Retocar el piso epóxico en el área de trabajo."
        },
        {
          "id": "p4_i3",
          "en": "Clean cable travel paths.",
          "es": "Limpiar las rutas de paso de cables."
        },
        {
          "id": "p4_i4",
          "en": "Mop and wax where required by the subcontract.",
          "es": "Trapear y encerar donde lo requiera el subcontrato."
        },
        {
          "id": "p4_i5",
          "en": "Verify floor repairs are acceptable and do not create trip hazards.",
          "es": "Verificar que las reparaciones del piso sean aceptables y no generen riesgos de tropiezo."
        }
      ]
    },
    {
      "id": "phase5",
      "number": 5,
      "name_en": "Final Cleaning and Project Closeout",
      "name_es": "Limpieza Final y Cierre del Proyecto",
      "items": [
        {
          "id": "p5_i1",
          "en": "Perform final cleaning.",
          "es": "Realizar la limpieza final."
        },
        {
          "id": "p5_i2",
          "en": "Complete all remaining punch list items.",
          "es": "Completar todos los elementos restantes de la lista de pendientes (punch list)."
        },
        {
          "id": "p5_i3",
          "en": "Verify all work areas are clean and ready for final acceptance.",
          "es": "Verificar que todas las áreas de trabajo estén limpias y listas para la aceptación final."
        },
        {
          "id": "p5_i4",
          "en": "Confirm no debris, dust, unused materials, or visible damage remain.",
          "es": "Confirmar que no queden escombros, polvo, materiales sin usar ni daños visibles."
        },
        {
          "id": "p5_i5",
          "en": "Prepare final inspection record.",
          "es": "Preparar el registro final de inspección."
        }
      ]
    }
  ]
};

window.CHECKLIST_DATA = CHECKLIST_DATA;
