import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  createDOMRenderer,
  RendererProvider,
  Button,
  Input,
  Text,
  tokens,
  Theme,
} from '@fluentui/react-components';
import { Settings24Regular, ShieldLock24Regular } from '@fluentui/react-icons';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'SmartPermissionsWebPartStrings';
import { SharePointService, isDebugEnabled } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { HomeView } from './HomeView';
import { PermissionsReportView } from './PermissionsReportView';
import { PermissionsExplorerView } from './PermissionsExplorerView';
import { UserAccessView } from './UserAccessView';
import { SettingsView } from './SettingsView';
import { readableTextOn } from '../utils/colorUtils';
import { setDisplayLocale } from '../utils/localeUtils';
import { useConfirmDialog } from './shared/ConfirmDialog';

export type AppView = 'home' | 'report' | 'explorer' | 'userAccess' | 'settings';

const LS_CONCURRENCY      = 'sp-smart-perms-concurrency';
const LS_GROUP_CAP        = 'sp-smart-perms-groupCap';
const LS_HIDDEN           = 'sp-smart-perms-includeHidden';
const LS_LIMITED_ACCESS   = 'sp-smart-perms-excludeLimitedAccess';
const LS_SHARING_LINKS    = 'sp-smart-perms-excludeSharingLinks';

const CONCURRENCY_DEFAULT = 4;
const CONCURRENCY_MIN = 1;
const CONCURRENCY_MAX = 10;
const GROUP_CAP_DEFAULT = 500;
const GROUP_CAP_MIN = 50;
const GROUP_CAP_MAX = 5000;

// A corrupt or cleared setting (NaN, 0, negative) must never reach the scan queue —
// TaskQueue never starts a task with a non-positive concurrency, so drain() would hang.
function clampInt(value: number, min: number, max: number, fallback: number): number {
  if (!isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, Math.round(value)));
}

function readIntSetting(key: string, min: number, max: number, fallback: number): number {
  try {
    return clampInt(parseInt(localStorage.getItem(key) ?? '', 10), min, max, fallback);
  } catch {
    return fallback;
  }
}

// Accepts a pasted page or library URL and reduces it to the web's base URL.
// Returns null when the input isn't a usable http(s) URL.
function normalizeSiteUrl(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;
  let url: URL;
  try {
    url = new URL(/^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`);
  } catch {
    return null;
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') return null;
  let path = url.pathname.replace(/\/+$/, '');
  // Drop a pasted page/view path (…/SitePages/Home.aspx, …/_layouts/15/viewlsts.aspx)
  path = path.replace(/\/(SitePages|_layouts|_catalogs|Lists|Shared%20Documents|Shared Documents)(\/.*)?$/i, '');
  path = path.replace(/\/[^/]+\.aspx$/i, '');
  return `${url.origin}${path}`;
}

export interface IBrandColors {
  primary: string;
  darkAlt: string;
  dark: string;
  darker: string;
  light: string;
  lighter: string;
}

function buildTheme(b: IBrandColors, isInverted: boolean): Theme {
  // Dark and "strong" section variants are inverted: start from the dark base theme and
  // use the palette's light shades wherever a brand color is drawn as text, a link or an
  // outline — the regular primary is tuned for light backgrounds and loses contrast there.
  const fg = isInverted ? b.light : b.primary;
  const fgHover = isInverted ? b.lighter : b.darkAlt;
  const fgPressed = isInverted ? b.lighter : b.dark;
  return {
    ...(isInverted ? webDarkTheme : webLightTheme),
    // Text on brand-colored fills (primary buttons) — white fails on light brand colors
    colorNeutralForegroundOnBrand:           readableTextOn(b.primary),
    // Background tokens (primary buttons, selected states)
    colorBrandBackground:                    b.primary,
    colorBrandBackgroundHover:               b.darkAlt,
    colorBrandBackgroundPressed:             b.dark,
    colorBrandBackgroundSelected:            b.darkAlt,
    colorBrandBackgroundStatic:              b.primary,
    colorBrandBackground2:                   b.lighter,
    colorBrandBackground2Hover:              b.light,
    colorBrandBackground2Pressed:            b.light,
    colorBrandBackground3Static:             b.dark,
    colorBrandBackground4Static:             b.darker,
    // Compound brand (checkboxes, radio buttons, sliders)
    colorCompoundBrandBackground:            b.primary,
    colorCompoundBrandBackgroundHover:       b.darkAlt,
    colorCompoundBrandBackgroundPressed:     b.dark,
    // Foreground tokens (icons, text, checkmarks)
    colorBrandForeground1:                   fg,
    colorBrandForeground2:                   fgHover,
    colorBrandForeground2Hover:              fgPressed,
    colorBrandForeground2Pressed:            b.darker,
    colorCompoundBrandForeground1:           fg,
    colorCompoundBrandForeground1Hover:      fgHover,
    colorCompoundBrandForeground1Pressed:    fgPressed,
    // Link foreground tokens
    colorBrandForegroundLink:                fg,
    colorBrandForegroundLinkHover:           fgHover,
    colorBrandForegroundLinkPressed:         fgPressed,
    colorBrandForegroundLinkSelected:        fg,
    // Stroke tokens (focus rings, borders)
    colorBrandStroke1:                       fg,
    colorBrandStroke2:                       b.light,
    colorBrandStroke2Hover:                  b.primary,
    colorBrandStroke2Pressed:                b.darkAlt,
    colorCompoundBrandStroke:                fg,
    colorCompoundBrandStrokeHover:           b.darkAlt,
    colorCompoundBrandStrokePressed:         b.dark,
  };
}

export interface AppProps {
  context: WebPartContext;
  sp: SharePointService;
  excel: ExcelExportService;
  defaultView?: AppView;
  brandColors: IBrandColors;
  /** True when the host section uses an inverted (dark/strong) theme variant. */
  isInverted?: boolean;
  /** Tools the page author hid from Home (and that cannot be opened). */
  hiddenViews?: AppView[];
  /** Site to open on instead of the page's own site. */
  defaultSiteUrl?: string;
  /** Hides Change URL so viewers can't redirect the part to another site. */
  lockSiteUrl?: boolean;
}

class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { error: Error | null }
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error: Error): { error: Error } {
    return { error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('[SmartPermissions] Render error:', error, info.componentStack);
  }

  render(): React.ReactNode {
    const { error } = this.state;
    if (error) {
      return (
        <div style={{
          padding: '16px', fontFamily: 'Consolas, monospace', fontSize: '13px',
          background: '#fff3f3', color: '#242424', border: '1px solid #c00', borderRadius: '4px', margin: '8px',
        }}>
          <strong style={{ color: '#c00', fontSize: '14px' }}>{strings.RenderErrorTitle}</strong>
          <br /><br />
          <strong>{strings.RenderErrorMessageLabel}</strong> {error.message || String(error)}
          <br /><br />
          {isDebugEnabled() && (
            <>
              <strong>{strings.RenderErrorStackLabel}</strong>
              <pre style={{
                fontSize: '11px', whiteSpace: 'pre-wrap', overflowWrap: 'anywhere',
                background: '#f5f5f5', padding: '8px', margin: '4px 0', borderRadius: '2px',
              }}>
                {error.stack ?? strings.RenderErrorNoStack}
              </pre>
            </>
          )}
        </div>
      );
    }
    return this.props.children;
  }
}

let renderer: ReturnType<typeof createDOMRenderer>;
try {
  renderer = createDOMRenderer(document);
} catch (e: any) {
  console.error('[SmartPermissions] createDOMRenderer failed:', e);
  throw e;
}

export const App: React.FC<AppProps> = ({ context, sp, excel, defaultView, brandColors, isInverted, hiddenViews, defaultSiteUrl, lockSiteUrl }) => {
  // Depend on the whole brandColors object, not just .primary — buildTheme reads
  // all six palette fields (dark/darkAlt/darker/light/lighter too), so a theme
  // change that alters those while primary stays constant previously left
  // hover/pressed/border brand tokens stale.
  const theme = React.useMemo(() => buildTheme(brandColors, !!isInverted), [brandColors, isInverted]);
  // Dates and numbers follow the SharePoint page language, not the browser's.
  setDisplayLocale(context.pageContext.cultureInfo?.currentUICultureName);
  const { confirm, dialog: confirmDialog } = useConfirmDialog();
  const onBrand = React.useMemo(() => readableTextOn(brandColors.primary), [brandColors.primary]);

  const [view, setView] = React.useState<AppView>(defaultView ?? 'home');

  // Switching screens replaces the content under the keyboard focus; without moving
  // focus, screen-reader and keyboard users are left on a control that no longer exists.
  const contentRef = React.useRef<HTMLDivElement>(null);
  const firstView = React.useRef(true);
  React.useEffect(() => {
    if (firstView.current) { firstView.current = false; return; }
    contentRef.current?.focus();
  }, [view]);
  const [prevView, setPrevView] = React.useState<AppView>('home');

  // The property pane can change defaultView while the part is mounted; useState only
  // reads its initial value, so follow later changes explicitly.
  const firstDefaultView = React.useRef(true);
  React.useEffect(() => {
    if (firstDefaultView.current) { firstDefaultView.current = false; return; }
    setView(defaultView ?? 'home');
  }, [defaultView]);

  // Always start from the current site — avoids stale cross-site URL from a previous session
  const defaultUrl = (defaultSiteUrl && normalizeSiteUrl(defaultSiteUrl)) || context.pageContext.web.absoluteUrl;
  const [siteUrl, setSiteUrl] = React.useState(() => defaultUrl);
  // The property pane can change the configured site while the part is mounted.
  const firstDefaultUrl = React.useRef(true);
  React.useEffect(() => {
    if (firstDefaultUrl.current) { firstDefaultUrl.current = false; return; }
    setSiteUrl(defaultUrl);
    setEditUrl(defaultUrl);
    setIsEditing(false);
  }, [defaultUrl]);
  const [editUrl, setEditUrl] = React.useState(siteUrl);
  const [isEditing, setIsEditing] = React.useState(false);

  // Persist settings to localStorage and keep service in sync
  const [includeHidden, setIncludeHidden] = React.useState(
    () => localStorage.getItem(LS_HIDDEN) === 'true',
  );
  const [scanConcurrency, setScanConcurrencyRaw] = React.useState(
    () => readIntSetting(LS_CONCURRENCY, CONCURRENCY_MIN, CONCURRENCY_MAX, CONCURRENCY_DEFAULT),
  );
  const [groupMemberCap, setGroupMemberCapRaw] = React.useState(
    () => readIntSetting(LS_GROUP_CAP, GROUP_CAP_MIN, GROUP_CAP_MAX, GROUP_CAP_DEFAULT),
  );
  const setScanConcurrency = (value: number): void =>
    setScanConcurrencyRaw(clampInt(value, CONCURRENCY_MIN, CONCURRENCY_MAX, CONCURRENCY_DEFAULT));
  const setGroupMemberCap = (value: number): void =>
    setGroupMemberCapRaw(clampInt(value, GROUP_CAP_MIN, GROUP_CAP_MAX, GROUP_CAP_DEFAULT));
  const [excludeLimitedAccess, setExcludeLimitedAccess] = React.useState(
    () => localStorage.getItem(LS_LIMITED_ACCESS) === 'true',
  );
  const [excludeSharingLinks, setExcludeSharingLinks] = React.useState(
    () => localStorage.getItem(LS_SHARING_LINKS) === 'true',
  );

  React.useEffect(() => { localStorage.setItem(LS_HIDDEN, String(includeHidden)); }, [includeHidden]);
  React.useEffect(() => { localStorage.setItem(LS_LIMITED_ACCESS, String(excludeLimitedAccess)); }, [excludeLimitedAccess]);
  React.useEffect(() => { localStorage.setItem(LS_SHARING_LINKS, String(excludeSharingLinks)); }, [excludeSharingLinks]);
  React.useEffect(() => {
    localStorage.setItem(LS_CONCURRENCY, String(scanConcurrency));
    sp.scanConcurrency = scanConcurrency;
  }, [scanConcurrency]);
  React.useEffect(() => {
    localStorage.setItem(LS_GROUP_CAP, String(groupMemberCap));
    sp.groupMemberCap = groupMemberCap;
  }, [groupMemberCap]);

  const [canManagePermissions, setCanManagePermissions] = React.useState<boolean | null>(null);
  const [siteOwners, setSiteOwners] = React.useState<{ title: string; email: string }[]>([]);

  React.useEffect(() => {
    setCanManagePermissions(null);
    let cancelled = false;
    sp.checkCanManagePermissions(siteUrl).then((can: boolean) => {
      if (cancelled) return;
      setCanManagePermissions(can);
      if (!can) sp.getSiteOwners(siteUrl).then(owners => { if (!cancelled) setSiteOwners(owners); }).catch(() => { /* ignore */ });
    }).catch(() => {
      // Fail closed, matching checkCanManagePermissions' own internal fallback —
      // don't grant Owner-level UI to a caller we couldn't verify.
      if (!cancelled) setCanManagePermissions(false);
    });
    return () => { cancelled = true; };
  }, [siteUrl]);

  // Home disables the tool cards for non-owners, but a tool can also be reached through
  // the defaultView property or by switching sites with Change URL. Send those users to
  // Home, where the "Site Owner required" explanation lives, instead of letting them hit
  // permission errors inside the tool.
  React.useEffect(() => {
    if (canManagePermissions === false && (view === 'report' || view === 'explorer' || view === 'userAccess')) {
      setView('home');
    }
  }, [canManagePermissions, view]);

  // A tool the page author hid must not be reachable through defaultView either.
  React.useEffect(() => {
    if (hiddenViews?.indexOf(view) !== undefined && hiddenViews.indexOf(view) !== -1) setView('home');
  }, [hiddenViews, view]);

  // Pre-fill login for the User Access view when launched from Explorer
  const [userAccessPrefill, setUserAccessPrefill] = React.useState<string | undefined>();

  // Tracks whether a scanning view reports a scan in progress (set via callback)
  const [scanBusy, setScanBusy] = React.useState(false);

  // Leaving a scanning view (Settings, a different site) unmounts it and aborts the scan,
  // so ask first, like the views' own Back buttons do.
  const runUnlessBusy = (action: () => void): void => {
    if (!scanBusy) { action(); return; }
    confirm({
      title: strings.ScanInProgressLabel,
      body: strings.ScanInProgressConfirm,
      confirmLabel: strings.LeaveButton,
      cancelLabel: strings.StayButton,
      onConfirm: action,
    });
  };

  const handleConnect = (): void => {
    const normalized = normalizeSiteUrl(editUrl);
    if (normalized && normalized !== siteUrl) {
      runUnlessBusy(() => {
        setSiteUrl(normalized);
        setEditUrl(normalized);
        setIsEditing(false);
      });
      return;
    }
    setEditUrl(normalized ?? siteUrl);
    setIsEditing(false);
  };

  const handleStartEdit = (): void => {
    setEditUrl(siteUrl);
    setIsEditing(true);
  };

  const handleCancelEdit = (): void => {
    setEditUrl(siteUrl);
    setIsEditing(false);
  };

  const handleOpenSettings = (): void => {
    runUnlessBusy(() => {
      setPrevView(view === 'settings' ? prevView : view);
      setView('settings');
    });
  };

  const handleNavigateToUserAccess = (loginName: string): void => {
    setUserAccessPrefill(loginName);
    setView('userAccess');
  };

  return (
    <ErrorBoundary>
    <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider
      theme={theme}
      dir={context.pageContext.cultureInfo?.isRightToLeft ? 'rtl' : 'ltr'}
      style={{ minHeight: '400px', position: 'relative' }}
    >
      {confirmDialog}
      <div ref={contentRef} tabIndex={-1} style={{ outline: 'none' }}>

      {/* Banner with site context and Change URL — shown on every view except Settings */}
      {view !== 'settings' && (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'auto 1fr auto',
            alignItems: 'center',
            paddingTop: tokens.spacingVerticalS,
            paddingBottom: tokens.spacingVerticalS,
            paddingLeft: tokens.spacingHorizontalM,
            paddingRight: tokens.spacingHorizontalS,
            background: brandColors.primary,
            gap: tokens.spacingHorizontalM,
          }}
        >
          {/* Left: branding */}
          <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, flexShrink: 0 }}>
            <ShieldLock24Regular style={{ color: onBrand, fontSize: '20px' }} />
            <Text style={{ color: onBrand, fontWeight: tokens.fontWeightSemibold, whiteSpace: 'nowrap' }}>
              {strings.AppTitle}
            </Text>
          </div>

          {/* Center: URL or edit input */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: tokens.spacingHorizontalS, minWidth: 0, overflow: 'hidden' }}>
            {isEditing ? (
              <>
                <Input
                  value={editUrl}
                  onChange={(_, d) => setEditUrl(d.value)}
                  placeholder={strings.SiteUrlPlaceholder}
                  style={{ minWidth: '200px', maxWidth: '400px', flexGrow: 1 }}
                  onKeyDown={(e) => { if (e.key === 'Enter') handleConnect(); }}
                />
                <Button appearance="secondary" onClick={handleConnect} disabled={!editUrl.trim()}>
                  {strings.ConnectButton}
                </Button>
                <Button appearance="transparent" style={{ color: onBrand, flexShrink: 0 }} onClick={handleCancelEdit}>
                  {strings.CancelButton}
                </Button>
              </>
            ) : (
              <>
                <span
                  style={{
                    width: '8px',
                    height: '8px',
                    borderRadius: '50%',
                    background: onBrand,
                    opacity: 0.75,
                    flexShrink: 0,
                    display: 'inline-block',
                  }}
                />
                <Text
                  style={{
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                    color: onBrand,
                  }}
                >
                  {siteUrl}
                </Text>
                {!lockSiteUrl && (
                  <Button
                    appearance="transparent"
                    size="small"
                    style={{ color: onBrand, flexShrink: 0 }}
                    onClick={handleStartEdit}
                  >
                    {strings.ChangeUrlButton}
                  </Button>
                )}
              </>
            )}
          </div>

          {/* Right: settings */}
          <Button
            appearance="transparent"
            icon={<Settings24Regular style={{ color: onBrand }} />}
            aria-label={strings.SettingsLabel}
            title={strings.SettingsLabel}
            onClick={handleOpenSettings}
          />
        </div>
      )}

      {view === 'home' && (
        <HomeView
          onNavigate={setView}
          canManagePermissions={canManagePermissions}
          hiddenViews={hiddenViews}
        />
      )}
      {view === 'report' && (
        <PermissionsReportView
          key={siteUrl}
          sp={sp}
          excel={excel}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          excludeLimitedAccess={excludeLimitedAccess}
          excludeSharingLinks={excludeSharingLinks}
          onBack={() => setView('home')}
          onBusyChange={setScanBusy}
        />
      )}
      {view === 'explorer' && (
        <PermissionsExplorerView
          key={siteUrl + String(includeHidden)}
          sp={sp}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          excludeLimitedAccess={excludeLimitedAccess}
          excludeSharingLinks={excludeSharingLinks}
          onBack={() => setView('home')}
          onNavigateToUserAccess={handleNavigateToUserAccess}
        />
      )}
      {view === 'userAccess' && (
        <UserAccessView
          key={siteUrl + String(includeHidden)}
          sp={sp}
          excel={excel}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          excludeLimitedAccess={excludeLimitedAccess}
          excludeSharingLinks={excludeSharingLinks}
          prefillLogin={userAccessPrefill}
          onPrefillUsed={() => setUserAccessPrefill(undefined)}
          onBack={() => setView('home')}
          onBusyChange={setScanBusy}
        />
      )}
      {view === 'settings' && (
        <SettingsView
          includeHidden={includeHidden}
          onIncludeHiddenChange={setIncludeHidden}
          excludeLimitedAccess={excludeLimitedAccess}
          onExcludeLimitedAccessChange={setExcludeLimitedAccess}
          excludeSharingLinks={excludeSharingLinks}
          onExcludeSharingLinksChange={setExcludeSharingLinks}
          scanConcurrency={scanConcurrency}
          onScanConcurrencyChange={setScanConcurrency}
          groupMemberCap={groupMemberCap}
          onGroupMemberCapChange={setGroupMemberCap}
          onBack={() => setView(prevView)}
        />
      )}

      </div>
    </FluentProvider>
    </RendererProvider>
    </ErrorBoundary>
  );
};
