// Skin icon names map to Fluent UI files: "lightbulb" → "lightbulb_20_regular.svg".
export function iconFile(name) {
  return /_(regular|filled)$/.test(name) ? `${name}.svg` : `${name}_20_regular.svg`;
}
