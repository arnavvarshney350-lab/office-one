// Builds "Office One.app" and a drag-to-Applications .dmg for macOS (Apple silicon by default).
//   npm run desktop:build            -> release/Office One-1.0.0-arm64.dmg and release/Office One-darwin-arm64/Office One.app
//   npm run desktop:build -- x64     -> Intel build
import { packager } from '@electron/packager'
import { execFileSync } from 'node:child_process'
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, symlinkSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import path from 'node:path'

const root = fileURLToPath(new URL('..', import.meta.url))
const arch = process.argv[2] === 'x64' ? 'x64' : 'arm64'
const { version } = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'))
const run = (cmd, args, opts = {}) => execFileSync(cmd, args, { stdio: 'inherit', cwd: root, ...opts })

console.log('1/4  Building the web app…')
run('npm', ['run', 'build'])

console.log('2/4  Staging the desktop shell…')
const stage = path.join(root, 'release', 'stage')
rmSync(stage, { recursive: true, force: true }); mkdirSync(stage, { recursive: true })
cpSync(path.join(root, 'desktop', 'main.cjs'), path.join(stage, 'main.cjs'))
cpSync(path.join(root, 'desktop', 'package.json'), path.join(stage, 'package.json'))
cpSync(path.join(root, 'dist'), path.join(stage, 'dist'), { recursive: true })
rmSync(path.join(stage, 'dist', 'sw.js'), { force: true }) // the offline worker is for browsers; the app already ships its files

console.log(`3/4  Packaging Office One.app (${arch})…`)
const out = path.join(root, 'release')
const [appPath] = await packager({
  dir: stage, out, name: 'Office One', appBundleId: 'com.officeone.app', appVersion: version, appCategoryType: 'public.app-category.productivity',
  platform: 'darwin', arch, icon: path.join(root, 'desktop', 'icon') /* extension is added per platform */, overwrite: true, prune: false,
  electronVersion: JSON.parse(readFileSync(path.join(root, 'node_modules', 'electron', 'package.json'), 'utf8')).version,
  extendInfo: { NSHumanReadableCopyright: 'Office One. Independent project, not affiliated with Microsoft.' },
  derefSymlinks: true,
})
const app = path.join(appPath, 'Office One.app')
// Re-sign ad hoc after the bundle was edited, so macOS runs it on Apple silicon without a developer certificate.
run('codesign', ['--force', '--deep', '--sign', '-', app])
run('codesign', ['--verify', '--deep', '--strict', app])

console.log('4/4  Creating the disk image…')
const dmgDir = path.join(out, 'dmg-staging')
rmSync(dmgDir, { recursive: true, force: true }); mkdirSync(dmgDir, { recursive: true })
run('ditto', [app, path.join(dmgDir, 'Office One.app')])
symlinkSync('/Applications', path.join(dmgDir, 'Applications'))
const dmg = path.join(out, `Office One-${version}-${arch}.dmg`)
rmSync(dmg, { force: true })
run('hdiutil', ['create', '-volname', 'Office One', '-srcfolder', dmgDir, '-ov', '-format', 'UDZO', '-fs', 'HFS+', dmg])
rmSync(dmgDir, { recursive: true, force: true }); rmSync(stage, { recursive: true, force: true })
if (!existsSync(dmg)) throw new Error('The disk image was not created')
console.log(`\nDone.\n  App: ${app}\n  DMG: ${dmg}`)
