// What the page list says about each page (pageList.ts): how many of its fields still need a value and how many
// proposals are still to check, the two kinds the host's "open" count adds up. Pure, no DOM.
import { pageIndexAtY, type ThumbPage } from '../apps/docs/src/renderer/components/NavPageThumbs'
import type { MarkerTone } from './embed-protocol'

export type PageCounts = { fields: number; missing: number; toReview: number }

/** a field and where its first line is drawn, in the same unzoomed canvas px as the pages */
export type PlacedField = { tone: MarkerTone; y: number }

const MISSING_TONES: MarkerTone[] = ['gap', 'conditional', 'refused']

/** the page canvas y is on; the band between two papers counts to the page below it */
const pageAt = (pages: ThumbPage[], y: number) => {
  const index = pageIndexAtY(pages, y)
  const end = pages[index].top + pages[index].height
  return index < pages.length - 1 && y >= end ? index + 1 : index
}

/** each page's counts; a field counts on the page its first line is on */
export function countByPage(pages: ThumbPage[], fields: PlacedField[]): PageCounts[] {
  const counts = pages.map(() => ({ fields: 0, missing: 0, toReview: 0 }))
  for (const { tone, y } of fields) {
    const page = counts[pageAt(pages, y)]
    page.fields += 1
    if (MISSING_TONES.includes(tone)) page.missing += 1
    if (tone === 'proposed') page.toReview += 1
  }
  return counts
}

type Labels = {
  heading: string
  list: string
  page: (n: number) => string
  missing: (n: number) => string
  toReview: (n: number) => string
  noFields: string
  done: string
}

const ENGLISH: Labels = {
  heading: 'Pages',
  list: 'Document pages',
  page: (n) => `Page ${n}`,
  missing: (n) => `${n} missing`,
  toReview: (n) => `${n} to review`,
  noFields: 'No input needed',
  done: 'Fully filled in',
}

const LABELS: Record<string, Labels> = {
  de: {
    heading: 'Seiten',
    list: 'Dokumentseiten',
    page: (n) => `Seite ${n}`,
    missing: (n) => (n === 1 ? `${n} fehlt` : `${n} fehlen`),
    toReview: (n) => `${n} zu prüfen`,
    noFields: 'Keine Eingaben nötig',
    done: 'Vollständig ausgefüllt',
  },
  fr: {
    heading: 'Pages',
    list: 'Pages du document',
    page: (n) => `Page ${n}`,
    missing: (n) => (n === 1 ? `${n} manquant` : `${n} manquants`),
    toReview: (n) => `${n} à vérifier`,
    noFields: 'Aucune saisie nécessaire',
    done: 'Entièrement rempli',
  },
}

export const pageListLabels = (lang: string) => LABELS[lang] ?? ENGLISH

/** what a page still needs, in words: its tooltip */
export function pageTitle({ fields, missing, toReview }: PageCounts, labels: Labels) {
  if (fields === 0) return labels.noFields
  const parts = [missing > 0 ? labels.missing(missing) : '', toReview > 0 ? labels.toReview(toReview) : '']
  return parts.filter((part) => part !== '').join(' · ') || labels.done
}
