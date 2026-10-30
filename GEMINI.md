# Reglas de Proyecto: Emprendimiento

Bienvenido al espacio de trabajo de **Emprendimiento**. Este entorno cuenta con más de **200 Agent Skills** configuradas para potenciar el desarrollo de negocios, ventas B2B, marketing, análisis de mercado y automatización con **Gemini** y el ecosistema **Antigravity**.

---

## 🎯 Habilidades Instaladas (.agents/skills)

Todas las habilidades están organizadas bajo [`.agents/skills/`](file:///home/victor/emprendimiento/.agents/skills) siguiendo el estándar de descubrimiento de Antigravity:

1. **Ventas y Negociación (B2B & Outbound)**:
   - Prospección y secuencias de correo (`email-sequence`, `outbound-prospecting`, `follow-up-discipline`).
   - Cualificación y llamadas (`discovery`, `budget-extraction-qualification`, `active-listening`, `asking-effective-questions`).
   - Manejo de objeciones y cierre (`objection-handling`, `closing`, `pricing-negotiation`, `scarcity-urgency-calibration`).
   - Análisis y pipelines (`pipeline-management`, `win-loss-reason-extraction`, `deal-review-win-loss`).

2. **Crecimiento, Marketing y GTM (OpenDirectory)**:
   - Generación de contenido y marca (`brand-alchemy`, `linkedin-post-generator`, `cook-the-blog`, `noise2blog`).
   - Inteligencia de mercado y competidores (`competitor-pr-finder`, `company-radar`, `map-your-market`, `hackernews-intel`, `reddit-icp-monitor`).
   - Lanzamientos y validación (`producthunt-launch-kit`, `oss-launch-kit`, `outreach-sequence-builder`, `yc-intent-radar-skill`).

3. **Productividad, Documentos y Flujo de Trabajo (Anthropic Skills)**:
   - Procesamiento y generación de documentos: Word ([`docx`](file:///home/victor/emprendimiento/.agents/skills/docx)), Presentaciones ([`pptx`](file:///home/victor/emprendimiento/.agents/skills/pptx)), Tablas de cálculo ([`xlsx`](file:///home/victor/emprendimiento/.agents/skills/xlsx)), Lectura y extracción ([`pdf`](file:///home/victor/emprendimiento/.agents/skills/pdf)).
   - Pruebas y creación de herramientas ([`webapp-testing`](file:///home/victor/emprendimiento/.agents/skills/webapp-testing), [`skill-creator`](file:///home/victor/emprendimiento/.agents/skills/skill-creator), [`mcp-builder`](file:///home/victor/emprendimiento/.agents/skills/mcp-builder)).

4. **SEO & GEO Toolkit (AgriciDaniel / claude-seo adaptado a Antigravity)**:
   - Auditorías completas y técnicas (`seo-audit`, `seo-technical`, `seo-page`, `seo-speed`).
   - Optimización semántica y GEO/AI Overviews (`seo-geo`, `seo-agentic`, `seo-content`, `seo-cluster`).
   - Estructuración y datos (`seo-schema`, `seo-sitemap`, `seo-hreflang`, `seo-programmatic`).
   - Backlinks y competidores (`seo-backlinks`, `seo-competitor-pages`, `seo-drift`).
   - Runtime de herramientas localizado en [`tools/seo/scripts/claude-seo`](file:///home/victor/emprendimiento/tools/seo/scripts/claude-seo).

---

## 🚀 Cómo usar las Habilidades

No necesitas invocar comandos manuales complejos; gracias a la **revelación progresiva** (*progressive disclosure*), Gemini detectará automáticamente cuándo una solicitud tuya coincide con la descripción de alguna de estas habilidades y cargará el runbook o playbook correspondiente en el momento necesario.

Por ejemplo, puedes pedir directamente:
- *"Ayúdame a redactar una secuencia de correos en frío para directores de TI"* (activará `email-sequence` y `cold-email-verifier`).
- *"Genera una propuesta comercial o guion de manejo de objeciones por precio"* (activará `objection-handling` y `pricing-negotiation`).
- *"Diseña un pitch deck en PPTX o un documento Word con nuestro plan comercial"* (activará `pptx` / `docx`).
- *"Analiza nuestro ICP en Reddit o monitoriza menciones de la competencia"* (activará `reddit-icp-monitor` / `competitor-pr-finder`).
