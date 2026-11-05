import * as strings from 'SmartPermissionsWebPartStrings';
import { ObjectType } from '../../models/models';

// ObjectType values are stable identifiers (they're stored in history and used as keys),
// so the on-screen label is mapped separately and localized.
export function objectTypeLabel(type: ObjectType | string): string {
  switch (type) {
    case ObjectType.Site: return strings.ObjectTypeSite;
    case ObjectType.Library: return strings.ObjectTypeLibrary;
    case ObjectType.List: return strings.ObjectTypeList;
    case ObjectType.Folder: return strings.ObjectTypeFolder;
    case ObjectType.File: return strings.ObjectTypeFile;
    default: return String(type);
  }
}
