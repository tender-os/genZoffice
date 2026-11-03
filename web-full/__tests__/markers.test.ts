import { describe, expect, it } from 'vitest'
import type { Decoration } from '@tiptap/pm/view'
import type { SlotMarker } from '../embed-protocol'
import { markerDecorations } from '../markers'
import { cell, doc, para, protectedBlock, table } from './markerSchema'

// positions: paragraph 0 at 0..3, protected block 1 at 3..4, table at 4, row at 5, cell at 6..14
// holding paragraph 2 at 7..10 and paragraph 3 at 10..13
const sample = doc(para(0, 'a'), protectedBlock(1), table(cell(para(2, 'b'), para(3, 'c'))))
const marker = (id: string, paragraph: number, scope: SlotMarker['scope'] = 'paragraph', tone: SlotMarker['tone'] = 'proposed'): SlotMarker =>
  ({ id, paragraph, scope, tone, label: id, actions: [] })
// a node decoration keeps its DOM attrs on its (untyped) type; nothing public reads them without a view
const attrsOf = (decoration: Decoration) => Reflect.get(Reflect.get(decoration, 'type'), 'attrs')
const described = (decoration: Decoration) =>
  ({ from: decoration.from, to: decoration.to, spec: decoration.spec, attrs: attrsOf(decoration) })
const node = (from: number, to: number, slotIds: string[], className: string) =>
  ({ from, to, spec: { slotIds, isField: false }, attrs: { class: className, 'data-vound-slot': slotIds.join(' ') } })

describe('markerDecorations', () => {
  it('decorates the paragraph node with the index, a protected block too, and the cell for cell scope', () => {
    const { decorations, dropped } = markerDecorations(
      sample,
      [marker('p', 0, 'paragraph', 'proposed'), marker('x', 1, 'paragraph', 'conditional'), marker('c', 3, 'cell', 'gap')],
      null,
    )
    expect(dropped).toEqual([])
    expect(decorations.map(described)).toEqual([
      node(0, 3, ['p'], 'vound-slot vound-slot-proposed'),
      node(3, 4, ['x'], 'vound-slot vound-slot-conditional'),
      node(6, 14, ['c'], 'vound-slot vound-slot-gap'),
    ])
  })

  it('drops an index matching no node or two nodes', () => {
    const twice = doc(para(0, 'a'), para(5, 'first half'), para(5, 'second half'))
    const { decorations, dropped } = markerDecorations(twice, [marker('split', 5), marker('missing', 9), marker('kept', 0)], null)
    expect(dropped).toEqual(['split', 'missing'])
    expect(decorations.map((d) => d.spec.slotIds)).toEqual([['kept']])
  })

  it('drops a cell marker whose paragraph sits outside a table cell', () => {
    const { decorations, dropped } = markerDecorations(sample, [marker('outside', 0, 'cell')], null)
    expect(dropped).toEqual(['outside'])
    expect(decorations).toEqual([])
  })

  it('marks only the current id with vound-slot-current', () => {
    const { decorations } = markerDecorations(sample, [marker('a', 0, 'paragraph', 'accepted'), marker('b', 2, 'paragraph', 'refused')], 'b')
    expect(decorations.map(described)).toEqual([
      node(0, 3, ['a'], 'vound-slot vound-slot-accepted'),
      node(7, 10, ['b'], 'vound-slot vound-slot-refused vound-slot-current'),
    ])
  })

  it('gives slots sharing a paragraph one decoration with every id and the tone that needs attention most', () => {
    const shared = [marker('ort', 0, 'paragraph', 'accepted'), marker('datum', 0, 'paragraph', 'gap'), marker('other', 2)]
    const { decorations, dropped } = markerDecorations(sample, shared, null)
    expect(dropped).toEqual([])
    expect(decorations.map(described)).toEqual([
      node(0, 3, ['ort', 'datum'], 'vound-slot vound-slot-gap'),
      node(7, 10, ['other'], 'vound-slot vound-slot-proposed'),
    ])
    // either id makes the shared node current
    expect(attrsOf(markerDecorations(sample, shared, 'ort').decorations[0]).class).toBe('vound-slot vound-slot-gap vound-slot-current')
    expect(attrsOf(markerDecorations(sample, shared, 'datum').decorations[0]).class).toBe('vound-slot vound-slot-gap vound-slot-current')
  })

  it('ranks the tones rejected, gap, refused, conditional, proposed, accepted', () => {
    const order: SlotMarker['tone'][] = ['rejected', 'gap', 'refused', 'conditional', 'proposed', 'accepted']
    const shownTone = (tones: SlotMarker['tone'][]) =>
      attrsOf(markerDecorations(sample, tones.map((tone, i) => marker(`s${i}`, 0, 'paragraph', tone)), null).decorations[0]).class
    order.forEach((tone, i) => {
      // the tone wins over every tone after it, whichever order the slots arrive in
      const rest = order.slice(i + 1)
      expect(shownTone([...rest, tone])).toBe(`vound-slot vound-slot-${tone}`)
      expect(shownTone([tone, ...rest])).toBe(`vound-slot vound-slot-${tone}`)
    })
  })

  it('groups two cell-scoped slots of one cell onto the cell', () => {
    const { decorations } = markerDecorations(sample, [marker('c2', 2, 'cell', 'proposed'), marker('c3', 3, 'cell', 'refused')], null)
    expect(decorations.map(described)).toEqual([node(6, 14, ['c2', 'c3'], 'vound-slot vound-slot-refused')])
  })
})
