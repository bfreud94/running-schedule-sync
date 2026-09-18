function parseMilesFromCell(value) {
  if (value === null || value === undefined) return 0;
  const text = String(value).replace(/\u00A0/g, ' ').trim().toLowerCase();
  if (text === '' || text === 'rest' || text.startsWith('data [')) return 0;

  const miles = parseFloat(text);
  return isNaN(miles) ? 0 : miles;
}

function isRunningActivity(activity) {
  const activityType = String(activity.type || '').toLowerCase();
  const sportType = String(activity.sport_type || '').toLowerCase();
  return activityType.includes('run') || sportType.includes('run');
}

function parseWorkout(description) {
  const match = String(description || '').match(/(?:^|\r?\n)workout:\s*\r?\n([^\r\n]*)/i);
  return match ? match[1].trim() : '';
}

function parseSupplementalWorkoutSection(description) {
  const match = String(description || '').match(/supplemental workouts?:?\s*\r?\n([\s\S]*)/i);
  if (!match) return '';

  const sectionLines = [];
  for (const line of match[1].split(/\r?\n/)) {
    if (!line.trim()) break;
    sectionLines.push(line);
  }
  return sectionLines.join('\n');
}

function isWeightDetail(detail) {
  return /\d+(?:\.\d+)?\s*(?:lbs?|pounds?|kgs?|kilograms?)\b/i.test(detail) || /^bodyweight$/i.test(detail);
}

function parseSupplementalExerciseLine(line, category) {
  const match = String(line || '').trim().match(/^\d+\.\s*([^()]+?)(?:\s*\(([^)]*)\))?\s*$/);
  if (!match) return null;

  const workout = match[1].trim();
  const details = String(match[2] || '').trim();
  const detailParts = details ? details.split(',').map(part => part.trim()).filter(Boolean) : [];
  const setsAndRepsMatch = (detailParts.shift() || '').match(/^(\d+)\s*x\s*(.+)$/i);

  let repsHoldTime = setsAndRepsMatch ? setsAndRepsMatch[2].trim() : '';
  const inlineWeightMatch = repsHoldTime.match(/^(.*?)\s*@\s*(.+)$/);
  const inlineWeight = inlineWeightMatch && isWeightDetail(inlineWeightMatch[2]) ? inlineWeightMatch[2].trim() : '';
  if (inlineWeight) repsHoldTime = inlineWeightMatch[1].trim();

  const weightParts = [inlineWeight, ...detailParts.filter(isWeightDetail)].filter(Boolean);
  const noteParts = detailParts.filter(part => !isWeightDetail(part));

  return {
    category,
    workout,
    sets: setsAndRepsMatch ? setsAndRepsMatch[1] : '',
    repsHoldTime,
    weight: weightParts.length ? weightParts.join('\n') : 'N/A',
    notes: noteParts.join(', ')
  };
}

function parseSupplementalWorkoutDetails(description) {
  const section = parseSupplementalWorkoutSection(description);
  if (!section) return { categories: [], exercises: [] };

  const categories = [];
  const exercises = [];
  let currentCategory = '';
  for (const line of section.split(/\r?\n/)) {
    const trimmedLine = line.trim();
    const categoryMatch = trimmedLine.match(/^([^:]+):\s*$/);
    if (categoryMatch) {
      currentCategory = categoryMatch[1].trim();
      if (currentCategory && !categories.includes(currentCategory)) categories.push(currentCategory);
      continue;
    }

    const exercise = parseSupplementalExerciseLine(trimmedLine, currentCategory);
    if (exercise) {
      exercises.push(exercise);
      continue;
    }

    const areaMatch = trimmedLine.match(/^area:\s*(.+)$/i);
    const supplementalWorkout = (areaMatch ? areaMatch[1] : trimmedLine).replace(/^[-*\u2022]\s*/, '');
    if (supplementalWorkout) {
      currentCategory = supplementalWorkout;
      if (!categories.includes(supplementalWorkout)) categories.push(supplementalWorkout);
    }
  }

  return { categories, exercises };
}

function parseSupplementalWorkouts(description) {
  return parseSupplementalWorkoutDetails(description).categories;
}

function calculateDailySupplementalWorkouts(activities, targetMonday) {
  const dailySupplementalWorkouts = [[], [], [], [], [], [], []];

  activities.forEach(activity => {
    const dayOffset = getDayOffset(parseActivityDate(activity), targetMonday);
    if (dayOffset < 0 || dayOffset > 6) return;

    parseSupplementalWorkouts(activity.description).forEach(supplementalWorkout => {
      if (!dailySupplementalWorkouts[dayOffset].includes(supplementalWorkout)) {
        dailySupplementalWorkouts[dayOffset].push(supplementalWorkout);
      }
    });
  });

  return dailySupplementalWorkouts;
}

function calculateDailyWorkouts(activities, targetMonday) {
  const dailyWorkouts = [[], [], [], [], [], [], []];

  activities.filter(isRunningActivity).forEach(activity => {
    const dayOffset = getDayOffset(parseActivityDate(activity), targetMonday);
    const workout = parseWorkout(activity.description);
    if (dayOffset >= 0 && dayOffset <= 6 && workout) {
      dailyWorkouts[dayOffset].push(workout);
    }
  });

  return dailyWorkouts;
}

function calculateDailyRunMiles(activities, targetMonday) {
  const metersToMiles = 1 / 1609.344;
  const dailyMiles = [0, 0, 0, 0, 0, 0, 0];

  activities.filter(isRunningActivity).forEach(activity => {
    const dayOffset = getDayOffset(parseActivityDate(activity), targetMonday);
    if (dayOffset >= 0 && dayOffset <= 6) {
      dailyMiles[dayOffset] += activity.distance * metersToMiles;
    }
  });

  return dailyMiles;
}