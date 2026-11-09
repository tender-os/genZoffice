import { describe, expect, it } from 'vitest'
import type { Decoration } from '@tiptap/pm/view'
import type { SlotMarker } from '../embed-protocol'
import { markerDecorations } from '../markers'
import { cell, doc, para, protectedBlock, schema, table } from './markerSchema'

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
      node(0, 3, ['p'], 'vound-slot vound-slot-proposed vound-slot-other-below'),
      node(3, 4, ['x'], 'vound-slot vound-slot-conditional vound-slot-other-above'),
      node(6, 14, ['c'], 'vound-slot vound-slot-gap'),
    ])
  })

  it('drops an index matching no node or two nodes apart, and marks both halves of a split', () => {
    const twice = doc(para(0, 'a'), para(5, 'one place'), para(6, 'b'), para(5, 'other place'))
    const { decorations, dropped } = markerDecorations(twice, [marker('copied', 5), marker('missing', 9), marker('kept', 0)], null)
    expect(dropped).toEqual(['copied', 'missing'])
    expect(decorations.map((d) => d.spec.slotIds)).toEqual([['kept']])
    const split = doc(para(0, 'a'), para(5, 'first half'), para(5, 'second half'))
    expect(markerDecorations(split, [marker('split', 5)], null).decorations.map((d) => d.from)).toEqual([3, 15])
  })

  it('marks every paragraph of the block-level control the field paragraph sits in, and no other', () => {
    const member = (index: number, text: string, attrs: Record<string, unknown>) =>
      schema.nodes.docParagraph.create({ docxPara: index, ...attrs }, schema.text(text))
    const inGroup = { sdtShell: JSON.stringify({ group: 2 }) }
    const body = doc(para(0, 'a'), member(1, 'first', inGroup), member(2, 'added', inGroup), para(3, 'b'))
    expect(markerDecorations(body, [marker('long', 1)], null).decorations.map((d) => [d.from, d.spec.slotIds])).toEqual([
      [3, ['long']],
      [10, ['long']],
    ])
    const inCell = doc(table(cell(member(4, 'label', {}), member(5, 'first', { cellControl: 0 }), member(6, 'added', { cellControl: 0 }))))
    expect(markerDecorations(inCell, [marker('cell field', 5)], null).decorations.map((d) => d.spec.slotIds)).toEqual([
      ['cell field'],
      ['cell field'],
    ])
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

  it('ranks the tones rejected, gap, refused, conditional, proposed, accepted, otherParty', () => {
    const order: SlotMarker['tone'][] = ['rejected', 'gap', 'refused', 'conditional', 'proposed', 'accepted', 'otherParty']
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

  it('marks where another field starts right above or below a block, so adjacent fields show apart', () => {
    // paragraphs 0..5: fields a, b, b (one field over two paragraphs), plain, c, then a cell field d in two rows
    const blocks = doc(para(0, 'a'), para(1, 'b1'), para(2, 'b2'), para(3, 'plain'), para(4, 'c'))
    const { decorations } = markerDecorations(blocks, [marker('a', 0), marker('b', 1), marker('b', 2), marker('c', 4)], null)
    expect(decorations.map((d) => attrsOf(d).class)).toEqual([
      'vound-slot vound-slot-proposed vound-slot-other-below',
      'vound-slot vound-slot-proposed vound-slot-other-above vound-slot-continued',
      'vound-slot vound-slot-proposed',
      'vound-slot vound-slot-proposed',
    ])
    const cells = doc(table(cell(para(0, 'x'), para(1, 'y'))), table(cell(para(2, 'z'))))
    const { decorations: inCells } = markerDecorations(cells, [marker('x', 0), marker('y', 1), marker('z', 2, 'cell')], null)
    expect(inCells.map((d) => attrsOf(d).class)).toEqual([
      'vound-slot vound-slot-proposed vound-slot-other-below',
      'vound-slot vound-slot-proposed vound-slot-other-above',
      'vound-slot vound-slot-proposed',
    ])
  })

  it('marks a block whose field goes on in the block below, so only the last one draws the trailing check', () => {
    const blocks = doc(para(0, 'a1'), para(1, 'a2'), para(2, 'b'))
    const { decorations } = markerDecorations(blocks, [marker('a', 0), marker('a', 1), marker('b', 2)], null)
    expect(decorations.map((d) => attrsOf(d).class)).toEqual([
      'vound-slot vound-slot-proposed vound-slot-continued',
      'vound-slot vound-slot-proposed vound-slot-other-below',
      'vound-slot vound-slot-proposed vound-slot-other-above',
    ])
    const rows = (...ids: number[]) =>
      schema.nodes.docTable.create(null, ids.map((id) => schema.nodes.docTableRow.create(null, [cell(para(id * 2, 'label')), cell(para(id * 2 + 1, 'value'))])))
    const { decorations: inCells } = markerDecorations(doc(rows(0, 1, 2)), [marker('c', 1, 'cell'), marker('c', 3, 'cell'), marker('d', 5, 'cell')], null)
    expect(inCells.map((d) => attrsOf(d).class)).toEqual([
      'vound-slot vound-slot-proposed vound-slot-continued',
      'vound-slot vound-slot-proposed',
      'vound-slot vound-slot-proposed',
    ])
  })
})
