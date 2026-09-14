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

function parseSupplementalWorkouts(description) {
  const match = String(description || '').match(/supplemental workouts?:?\s*\r?\n([\s\S]*)/i);
  if (!match) return [];

  const supplementalWorkouts = [];
  for (const line of match[1].split(/\r?\n/)) {
    const supplementalWorkout = line.trim().replace(/^[-*\u2022]\s*/, '');
    if (!supplementalWorkout) break;

    supplementalWorkouts.push(supplementalWorkout);
  }
  return supplementalWorkouts;
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