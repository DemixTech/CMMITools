/**
 * Configuration loader for CAS Form Populator
 *
 * Reads:
 *   ../cas-project-config.json   (project paths + cas settings)
 *   ../.secrets/keys.json        (credentials)
 *
 * Paths are resolved relative to this file (cas-scraper/config.ts), so it
 * doesn't matter whether `npm run populate` is launched from cas-scraper/
 * or from somewhere else.
 */

import * as fs from 'fs';
import * as path from 'path';

const projectConfigPath = path.resolve(__dirname, '..', 'cas-project-config.json');
const keysPath          = path.resolve(__dirname, '..', '.secrets', 'keys.json');

let projectConfig: any = null;
let keysConfig:    any = null;

if (fs.existsSync(projectConfigPath)) {
  try {
    projectConfig = JSON.parse(fs.readFileSync(projectConfigPath, 'utf-8'));
  } catch (e) {
    console.warn(`⚠️  Failed to parse ${projectConfigPath}: ${(e as Error).message}`);
  }
} else {
  console.warn(`⚠️  Project config not found at ${projectConfigPath}`);
}

if (fs.existsSync(keysPath)) {
  try {
    keysConfig = JSON.parse(fs.readFileSync(keysPath, 'utf-8'));
  } catch (e) {
    console.warn(`⚠️  Failed to parse ${keysPath}: ${(e as Error).message}`);
  }
} else {
  console.warn(`⚠️  Keys file not found at ${keysPath}`);
}

export const CONFIG = {
  // Paths (exported so other modules / error messages can reference them)
  projectConfigPath,
  keysPath,

  casBaseUrl:        projectConfig?.cas?.baseUrl     || 'https://cas.cmmiinstitute.com',
  loginUrl:          projectConfig?.cas?.loginUrl    || 'https://cmmiinstitute.com/login',
  appraisalId:       projectConfig?.project?.casId   || '81846',

  // Continue from specific page (empty string = start from beginning)
  continueFromPage:  projectConfig?.cas?.continueFromPage || '',

  // Credentials from keys.json (preferred) or environment variables (fallback)
  email:             keysConfig?.cas?.email    || process.env.CAS_EMAIL    || '',
  password:          keysConfig?.cas?.password || process.env.CAS_PASSWORD || '',
  staySignedIn:      keysConfig?.cas?.staySignedIn?.toLowerCase() === 'yes',

  // Excel source file (from cas-project-config.json files.target)
  excelFile:         projectConfig?.files?.target || '',

  // Timing
  navigationTimeout: 90000,
  waitAfterAction:   1000,

  // Logging
  logFile:           'populate_log.json',
  htmlLogDir:        'html_logs',

  // Debug mode - if true, prompt for input after each page; if false, auto-continue
  debugMode:         projectConfig?.cas?.debugMode ?? true,

  // Auto exit on complete - if true, exit when reaching /sample-scope; if false, prompt
  autoExitOnComplete: projectConfig?.cas?.autoExitOnComplete ?? false,

  // Pages to skip (no processing needed)
  skipPages: [
    '/org-unit-project-appraisal-scope',
    '/include-project'
  ],

  // Final page - reaching this means Phase 1 is complete
  finalPage: '/sample-scope'
};
