export type CategoryTuple = [name: string, id: number]

export const CATEGORIES: CategoryTuple[] = [
  ['Accounts Billing', 252434],
  ['Construction Or Project Related', 277747],
  ['Design Related Issue', 277632],
  ['Estate Infra Outer Area from the plot', 267181],
  ['FM Common Area Related Issue', 277745],
  ['IT WIFI Network', 277816],
  ['Products Appliances', 277744],
]

export const CATEGORY_IDS = CATEGORIES.map(([, id]) => id)

/** Category name keyed by dashboard category id. */
export const CATEGORY_NAME = new Map(CATEGORIES.map(([name, id]) => [id, name]))
