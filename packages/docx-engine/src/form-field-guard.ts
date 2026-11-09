const FORM_FIELD_MARKS = ['FORMTEXT', 'FORMCHECKBOX', 'FORMDROPDOWN', '<w:sdt>', '<w:sdt ']

const occurrences = (xml: string, mark: string) => xml.split(mark).length - 1

/** Vound fork: whether rebuilt paragraph XML still holds every form field and content control of the original */
export const keepsFormFields = (original: string, rebuilt: string) =>
  FORM_FIELD_MARKS.every((mark) => occurrences(rebuilt, mark) >= occurrences(original, mark))
