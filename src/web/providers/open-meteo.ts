import { MAX_RESULT_CONTENT_CHARS, type SearchProviderAdapter, type SearchResult } from './index';

export const openMeteoProvider: SearchProviderAdapter = {
  id: 'open-meteo',
  label: 'Open-Meteo (keyless)',
  requiresKey: false,
  endpoint: 'https://api.open-meteo.com/v1/forecast',
  signupUrl: 'https://open-meteo.com/en/docs',

  async search(query, opts): Promise<SearchResult[]> {
    const coordinate = parseCoordinate(query);
    const results = coordinate
      ? coordinateResults(query, coordinate)
      : placeResults(query);
    return results.slice(0, opts.maxResults).map(item => ({
      ...item,
      ...(opts.includeContent ? { content: capContent(item.content) } : {}),
    }));
  },
};

interface Candidate {
  title: string;
  url: string;
  snippet: string;
  content: string;
}

function coordinateResults(query: string, coordinate: { latitude: string; longitude: string }): Candidate[] {
  const encodedLat = encodeURIComponent(coordinate.latitude);
  const encodedLon = encodeURIComponent(coordinate.longitude);
  const base = `https://api.open-meteo.com/v1/forecast?latitude=${encodedLat}&longitude=${encodedLon}`;
  const hourly = `${base}&hourly=temperature_2m,relative_humidity_2m,precipitation,wind_speed_10m&forecast_days=7`;
  const daily = `${base}&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,wind_speed_10m_max&timezone=auto`;
  return [
    candidate('Open-Meteo hourly forecast', hourly, `Hourly forecast for ${coordinate.latitude}, ${coordinate.longitude}.`, query),
    candidate('Open-Meteo daily forecast', daily, `Daily forecast for ${coordinate.latitude}, ${coordinate.longitude}.`, query),
  ];
}

function placeResults(query: string): Candidate[] {
  const place = query.trim();
  const encodedPlace = encodeURIComponent(place);
  const geocodeUrl = `https://geocoding-api.open-meteo.com/v1/search?name=${encodedPlace}&count=10&language=en&format=json`;
  return [
    candidate('Open-Meteo geocoding search', geocodeUrl, 'Find latitude/longitude for a place name before fetching weather.', query),
    candidate('Open-Meteo forecast API docs', 'https://open-meteo.com/en/docs', 'Forecast endpoint requires latitude and longitude; use the geocoding result to build a forecast URL.', query),
  ];
}

function candidate(title: string, url: string, snippet: string, query: string): Candidate {
  return {
    title,
    url,
    snippet,
    content: [`Query: ${query}`, `Title: ${title}`, `URL: ${url}`, `Note: ${snippet}`].join('\n'),
  };
}

function parseCoordinate(query: string): { latitude: string; longitude: string } | null {
  const match = query.match(/(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)/);
  if (!match) return null;
  const latitude = Number(match[1]);
  const longitude = Number(match[2]);
  if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) return null;
  if (latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) return null;
  return { latitude: match[1], longitude: match[2] };
}

function capContent(raw: string): string {
  return raw.length <= MAX_RESULT_CONTENT_CHARS ? raw : `${raw.slice(0, MAX_RESULT_CONTENT_CHARS)}... [truncated: API candidate metadata continues beyond this point]`;
}
