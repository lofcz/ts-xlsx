import {openElements} from '../../xml/xml-read.ts';
import {transitionalRelationshipType} from './strict-relationships.ts';

export interface RelationshipRecord {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

// OPC makes Id, Type and Target all mandatory, so an element missing any of the three is not a
// relationship and is skipped rather than half-read.
export function parseRelationshipRecords(xml: string): RelationshipRecord[] {
  const records: RelationshipRecord[] = [];
  for (const {attrs} of openElements(xml, 'Relationship')) {
    if (attrs.Id !== undefined && attrs.Type !== undefined && attrs.Target !== undefined) {
      records.push({
        id: attrs.Id,
        // In its Transitional spelling, so a Strict package's relationships are recognised by the same
        // names and a preserved one is written back in the spelling of the package it lands in.
        type: transitionalRelationshipType(attrs.Type),
        target: attrs.Target,
        external: attrs.TargetMode === 'External',
      });
    }
  }
  return records;
}
