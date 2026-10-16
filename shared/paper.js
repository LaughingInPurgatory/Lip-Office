export const PAPER_SIZES = {
  letter: { label: 'Letter', inches: [8.5, 11], twips: [12240, 15840] },
  legal: { label: 'Legal', inches: [8.5, 14], twips: [12240, 20160] },
  tabloid: { label: 'Tabloid', inches: [11, 17], twips: [15840, 24480] },
  a3: { label: 'A3', inches: [11.69, 16.54], twips: [16834, 23818] },
  a4: { label: 'A4', inches: [8.27, 11.69], twips: [11909, 16834] },
  a5: { label: 'A5', inches: [5.83, 8.27], twips: [8395, 11909] },
  b5: { label: 'B5', inches: [6.93, 9.84], twips: [9979, 14170] },
};

export function paperDimensions(size) {
  return PAPER_SIZES[size]?.inches || PAPER_SIZES.letter.inches;
}

export function paperTwips(size) {
  return PAPER_SIZES[size]?.twips || PAPER_SIZES.letter.twips;
}

export function paperSizeFromDimensions(width, height) {
  const [shortSide, longSide] = [Number(width), Number(height)].sort((a, b) => a - b);
  let closest = 'letter';
  let distance = Infinity;
  for (const [key, paper] of Object.entries(PAPER_SIZES)) {
    const [paperShort, paperLong] = [...paper.inches].sort((a, b) => a - b);
    const nextDistance = (shortSide - paperShort) ** 2 + (longSide - paperLong) ** 2;
    if (nextDistance < distance) {
      closest = key;
      distance = nextDistance;
    }
  }
  return closest;
}
