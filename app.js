const __rootScope = typeof globalThis !== "undefined"
  ? globalThis
  : (typeof window !== "undefined" ? window : (typeof self !== "undefined" ? self : {}));

if (typeof __rootScope.structuredClone !== "function") {
  __rootScope.structuredClone = (value) => JSON.parse(JSON.stringify(value));
}

if (!__rootScope.crypto) {
  __rootScope.crypto = {};
}
if (typeof __rootScope.crypto.randomUUID !== "function") {
  __rootScope.crypto.randomUUID = () => {
    const pattern = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx";
    return pattern.replace(/[xy]/g, (char) => {
      const random = Math.floor(Math.random() * 16);
      const value = char === "x" ? random : ((random & 0x3) | 0x8);
      return value.toString(16);
    });
  };
}

const STORAGE_KEY = "maitavern-state-v8"; // legacy fallback (migration only)
const USER_FILE_STORAGE_KEY = "maitavern-user-file-v1"; // legacy fallback (migration only)
const USER_FILE_SCHEMA_VERSION = 1;
const USER_FILE_NAME = "user";
const USER_FILE_ENDPOINT = "/api/user";
const USER_FILE_INDEXED_DB_NAME = "maitavern-user-db";
const USER_FILE_INDEXED_DB_STORE = "payloads";
const USER_FILE_INDEXED_DB_KEY = "user-file";

let userFileWriteTimer = null;
let pendingUserFilePayload = null;
let userFileWriteInFlight = false;
let deviceStorageUnavailable = false;
let userFileIndexedDbPromise = null;

const defaults = {
  hasEnteredApp: false,
  activeView: "homeView",
  endpoint: "",
  apiKey: "",
  lastGenerationType: "",
  authHeader: "Authorization",
  authPrefix: "Bearer ",
  extraHeaders: "",
  extraBody: "",
  models: ["gpt-4o-mini"],
  selectedModel: "gpt-4o-mini",
  systemPrompt: "You are a helpful assistant.",
  messages: [
    {
      role: "system",
      content: "Welcome to MaiTavern. Open settings to add your endpoint, API key, and models.",
    },
  ],
  characters: [],
  providers: [],
  assets: [],
  lorebooks: [],
  lorebookSettings: {
    scanDepth: 4,
    tokenBudget: 1200,
    recursiveScanning: true,
  },
  presets: [
    {
      id: "default-preset",
      name: "Default preset",
      systemPrompt: "You are a helpful assistant.",
      temperature: 0.7,
      topP: 1,
      maxTokens: 300,
      stream: false,
      contextTemplate: "{{#if system}}{{system}}\n{{/if}}{{#if description}}{{description}}\n{{/if}}{{#if personality}}{{personality}}\n{{/if}}{{#if scenario}}{{scenario}}\n{{/if}}{{#if persona}}{{persona}}\n{{/if}}{{trim}}",
      chatStart: "***",
      exampleSeparator: "***",
      instruct: {
        enabled: false,
        inputSequence: "<|im_start|>user",
        outputSequence: "<|im_start|>assistant",
        systemSequence: "<|im_start|>system",
        stopSequence: "<|im_end|>",
      },
      useSystemPrompt: true,
      namesBehavior: "",
      prompts: [],
      promptOrder: [],
      source: "maitavern",
    },
  ],
  activePresetId: "default-preset",
  activeProviderId: "",
  chats: [],
  recentChats: [],
  currentChatId: "",
  currentChatTitle: "New chat",
  promptLogs: [],
  macroGlobalVars: {},
  macroChatVars: {},
  customization: {
    chatStyle: "bubble",
    avatarSize: 40,
    defaultTextColor: "#f5f7fb",
    appFontSource: "system",
    appFontFamily: "Inter",
    appFontFileName: "",
    appFontDataUrl: "",
    appFontGoogleName: "",
    appFontGoogleCss: "",
    bubble: {
      assistantBubbleColor: "#111114",
      userBubbleColor: "#f1f4fb",
      showAssistantName: true,
      showUserName: true,
      assistantAvatarSize: 40,
      userAvatarSize: 40,
      assistantAvatarShape: "cutout",
      userAvatarShape: "cutout",
      showAssistantAvatar: true,
      showUserAvatar: true,
    },
    novel: {
      showAssistantName: true,
      showUserName: true,
      showAssistantAvatar: true,
      showUserAvatar: true,
      assistantAvatarSize: 64,
      userAvatarSize: 64,
      assistantAvatarShape: "cutout",
      userAvatarShape: "cutout",
      assistantAvatarAlign: "left",
      userAvatarAlign: "right",
      assistantTextFlow: "wrap",
      userTextFlow: "wrap",
      assistantBackgroundColor: "#111114",
      userBackgroundColor: "#151a24",
      assistantBackgroundImage: "",
      userBackgroundImage: "",
      assistantBackgroundBlur: 0,
      userBackgroundBlur: 0,
      chatBackgroundColor: "#000000",
      chatBackgroundImage: "",
      chatBackgroundBlur: 0,
    },
    wrapRules: [
      { id: "quotes-double", title: '"..."', start: '"', end: '"', color: "#7dd3fc", enabled: true },
      { id: "asterisk", title: "*...*", start: "*", end: "*", color: "#f9a8d4", enabled: true },
      { id: "guillemets", title: "«...»", start: "«", end: "»", color: "#c084fc", enabled: true },
    ],
  },
};

let state = structuredClone(defaults);
let currentController = null;
let presetEditorDraft = null;
let messageActionState = { index: -1, mode: "" };
let expandedMessageMenuIndex = -1;
let messageInlineEditIndex = -1;
let drawerPanelStack = ["root"];
let drawerLorebookSelectedId = "";
let wrapRuleSearchQuery = "";
let appFontStyleTag = null;
let lorebookEntryDraft = [];
let lorebookEntryLongPress = {
  timer: null,
  pointerId: null,
  sourceIndex: -1,
  targetIndex: -1,
  active: false,
  moveHandler: null,
  upHandler: null,
  autoScrollRaf: null,
  latestClientX: 0,
  latestClientY: 0,
  scrollContainer: null,
};
let lorebookEntrySearchQuery = "";
let libraryHubSearchQuery = "";
let drawerPresetEditorDraft = null;
let drawerPresetPromptDragIndex = -1;
let drawerSelectedLogId = "";


// Lorebook V2 runtime state
const LOREBOOK_V2_ENABLED = true;
const LOREBOOK_V2_COLORS = ["#5B8CFF", "#2EC4B6", "#A78BFA", "#FF7A59", "#F59E0B", "#22C55E", "#EC4899", "#06B6D4"];
const LOREBOOK_V2_SORT_OPTIONS = {
  ALPHABETICAL: "alphabetical",
  DATE_MODIFIED: "date_modified",
  DATE_CREATED: "date_created",
  ENTRY_COUNT: "entry_count",
  MANUAL: "manual",
  PRIORITY: "priority",
  TOKEN_COUNT: "token_count",
  RECENTLY_MODIFIED: "recently_modified",
};

const LBV2_INJECTION_POSITIONS = [
  { value: "before_system", label: "Before System Prompt" },
  { value: "after_system", label: "After System Prompt" },
  { value: "before_author_note", label: "Before Author's Note" },
  { value: "after_author_note", label: "After Author's Note" },
  { value: "top_chat", label: "Top of Chat History" },
  { value: "bottom_chat", label: "Bottom of Chat History" },
  { value: "depth", label: "@ Depth" },
];

let lbv2State = {
  level: 1,
  stackHistory: [1],
  activeLorebookId: "",
  activeEntryId: "",
  lorebookSearchQuery: "",
  entrySearchQuery: "",
  lorebookSortBy: LOREBOOK_V2_SORT_OPTIONS.DATE_MODIFIED,
  entrySortBy: LOREBOOK_V2_SORT_OPTIONS.MANUAL,
  entryFilter: "all",
  entryTagFilter: "",
  bulkMode: false,
  bulkSelectedEntryIds: new Set(),
  expandedEntrySwipeId: "",
  editingEntry: null,
  sheet: {
    key: "",
    payload: null,
  },
  dialog: {
    onConfirm: null,
  },
  saveIndicatorTimer: null,
  autosaveTimer: null,
  suppressEditorEvents: false,
  gesture: {
    pointerId: null,
    startX: 0,
    startY: 0,
    tracking: false,
  },
  drag: {
    pointerId: null,
    sourceId: "",
    targetId: "",
    active: false,
    timer: null,
  },
  stepperHold: {
    timer: null,
    interval: null,
  },
};

const els = {
  welcomeScreen: document.getElementById("welcomeScreen"),
  appShell: document.getElementById("appShell"),
  enterAppBtn: document.getElementById("enterAppBtn"),
  homeBrandBtn: document.getElementById("homeBrandBtn"),
  settingsToggle: document.getElementById("settingsToggle"),
  settingsPanel: document.getElementById("settingsPanel"),
  closeSettingsBtn: document.getElementById("closeSettingsBtn"),
  endpointInput: document.getElementById("endpointInput"),
  apiKeyInput: document.getElementById("apiKeyInput"),
  authHeaderInput: document.getElementById("authHeaderInput"),
  authPrefixInput: document.getElementById("authPrefixInput"),
  modelsInput: document.getElementById("modelsInput"),
  extraHeadersInput: document.getElementById("extraHeadersInput"),
  extraBodyInput: document.getElementById("extraBodyInput"),
  systemPromptInput: document.getElementById("systemPromptInput"),
  saveSettingsBtn: document.getElementById("saveSettingsBtn"),
  clearStorageBtn: document.getElementById("clearStorageBtn"),
  fetchModelsBtn: document.getElementById("fetchModelsBtn"),
  fetchModelsStatus: document.getElementById("fetchModelsStatus"),
  modelSelect: document.getElementById("modelSelect"),
  messages: document.getElementById("messages"),
  composer: document.getElementById("composer"),
  messageInput: document.getElementById("messageInput"),
  sendBtn: document.getElementById("sendBtn"),
  stopBtn: document.getElementById("stopBtn"),
  statusText: document.getElementById("statusText"),
  homeGreeting: document.getElementById("homeGreeting"),
  newChatBtn: document.getElementById("newChatBtn"),
  messageTemplate: document.getElementById("messageTemplate"),
  navCards: Array.from(document.querySelectorAll(".nav-card")),
  views: Array.from(document.querySelectorAll(".view")),
  libraryView: document.getElementById("libraryView"),
  characterList: document.getElementById("characterList"),
  characterSearchInput: document.getElementById("characterSearchInput"),
  characterStudioBtn: document.getElementById("characterStudioBtn"),
  characterStudioImportJsonBtn: document.getElementById("characterStudioImportJsonBtn"),
  characterStudioImportAvatarBtn: document.getElementById("characterStudioImportAvatarBtn"),
  characterImportInput: document.getElementById("characterImportInput"),
  characterAvatarImportInput: document.getElementById("characterAvatarImportInput"),
  characterEditorCard: document.getElementById("characterEditorCard"),
  characterNameInput: document.getElementById("characterNameInput"),
  characterScenarioInput: document.getElementById("characterScenarioInput"),
  characterPersonalityInput: document.getElementById("characterPersonalityInput"),
  characterGreetingInput: document.getElementById("characterGreetingInput"),
  characterAlternateGreetingsInput: document.getElementById("characterAlternateGreetingsInput"),
  characterAvatarInput: document.getElementById("characterAvatarInput"),
  characterAvatarPreview: document.getElementById("characterAvatarPreview"),
  characterSaveBtn: document.getElementById("characterSaveBtn"),
  characterCancelBtn: document.getElementById("characterCancelBtn"),
  lorebookScanDepthInput: document.getElementById("lorebookScanDepthInput"),
  lorebookTokenBudgetInput: document.getElementById("lorebookTokenBudgetInput"),
  lorebookRecursiveInput: document.getElementById("lorebookRecursiveInput"),
  saveLorebookSettingsBtn: document.getElementById("saveLorebookSettingsBtn"),
  lorebookSettingsStatus: document.getElementById("lorebookSettingsStatus"),
  lorebookSearchInput: document.getElementById("lorebookSearchInput"),
  lorebookImportBtn: document.getElementById("lorebookImportBtn"),
  lorebookExportBtn: document.getElementById("lorebookExportBtn"),
  lorebookNewBtn: document.getElementById("lorebookNewBtn"),
  lorebookImportInput: document.getElementById("lorebookImportInput"),
  libraryHubPanel: document.getElementById("libraryHubPanel"),
  libraryHubSearchInput: document.getElementById("libraryHubSearchInput"),
  libraryHubGrid: document.getElementById("libraryHubGrid"),
  libraryHubEmptyState: document.getElementById("libraryHubEmptyState"),
  libraryCategoryChips: Array.from(document.querySelectorAll(".library-category-chip")),
  libraryPanels: Array.from(document.querySelectorAll("#libraryView .library-panel")),
  libraryPanelBackBtns: Array.from(document.querySelectorAll(".library-panel-back-btn")),
  lorebookEditorView: document.getElementById("lorebookEditorView"),
  lorebookEditorBackBtn: document.getElementById("lorebookEditorBackBtn"),
  lorebookEditorImportBtn: document.getElementById("lorebookEditorImportBtn"),
  lorebookEditorImportInput: document.getElementById("lorebookEditorImportInput"),
  lorebookEditorStatus: document.getElementById("lorebookEditorStatus"),
  lorebookNameInput: document.getElementById("lorebookNameInput"),
  lorebookDescriptionInput: document.getElementById("lorebookDescriptionInput"),
  lorebookSaveBtn: document.getElementById("lorebookSaveBtn"),
  lorebookCancelBtn: document.getElementById("lorebookCancelBtn"),
  lorebookList: document.getElementById("lorebookList"),
  lorebookEntrySearchInput: document.getElementById("lorebookEntrySearchInput"),
  lorebookEntryNewBtn: document.getElementById("lorebookEntryNewBtn"),
  lorebookEntryList: document.getElementById("lorebookEntryList"),
  lorebookEntryEditorCard: document.getElementById("lorebookEntryEditorCard"),
  lorebookEntryTitleInput: document.getElementById("lorebookEntryTitleInput"),
  lorebookEntryKeysInput: document.getElementById("lorebookEntryKeysInput"),
  lorebookEntrySecondaryKeysInput: document.getElementById("lorebookEntrySecondaryKeysInput"),
  lorebookEntryContentInput: document.getElementById("lorebookEntryContentInput"),
  lorebookEntryPriorityInput: document.getElementById("lorebookEntryPriorityInput"),
  lorebookEntryProbabilityInput: document.getElementById("lorebookEntryProbabilityInput"),
  lorebookEntryEnabledInput: document.getElementById("lorebookEntryEnabledInput"),
  lorebookEntryConstantInput: document.getElementById("lorebookEntryConstantInput"),
  lorebookEntryCaseSensitiveInput: document.getElementById("lorebookEntryCaseSensitiveInput"),
  lorebookEntrySelectiveInput: document.getElementById("lorebookEntrySelectiveInput"),
  lorebookEntrySelectiveLogicInput: document.getElementById("lorebookEntrySelectiveLogicInput"),
  lorebookEntrySaveBtn: document.getElementById("lorebookEntrySaveBtn"),
  lorebookEntryDeleteBtn: document.getElementById("lorebookEntryDeleteBtn"),
  lorebookEntryCancelBtn: document.getElementById("lorebookEntryCancelBtn"),
  providerSearchInput: document.getElementById("providerSearchInput"),
  providerImportBtn: document.getElementById("providerImportBtn"),
  providerExportBtn: document.getElementById("providerExportBtn"),
  providerNewBtn: document.getElementById("providerNewBtn"),
  providerImportInput: document.getElementById("providerImportInput"),
  providerFormCard: document.getElementById("providerFormCard"),
  providerNameInput: document.getElementById("providerNameInput"),
  providerUrlInput: document.getElementById("providerUrlInput"),
  providerApiKeyInput: document.getElementById("providerApiKeyInput"),
  providerThinkStartInput: document.getElementById("providerThinkStartInput"),
  providerThinkEndInput: document.getElementById("providerThinkEndInput"),
  providerShowThinkingInput: document.getElementById("providerShowThinkingInput"),
  providerIncludeThinkingInput: document.getElementById("providerIncludeThinkingInput"),
  providerSaveBtn: document.getElementById("providerSaveBtn"),
  providerCancelBtn: document.getElementById("providerCancelBtn"),
  providerFetchModelsBtn: document.getElementById("providerFetchModelsBtn"),
  providerFetchModelsStatus: document.getElementById("providerFetchModelsStatus"),
  providerList: document.getElementById("providerList"),
  presetNewBtn: document.getElementById("presetNewBtn"),
  presetFormCard: document.getElementById("presetFormCard"),
  presetImportBtn: document.getElementById("presetImportBtn"),
  presetExportBtn: document.getElementById("presetExportBtn"),
  presetImportInput: document.getElementById("presetImportInput"),
  presetNameInput: document.getElementById("presetNameInput"),
  presetTemperatureInput: document.getElementById("presetTemperatureInput"),
  presetTopPInput: document.getElementById("presetTopPInput"),
  presetMaxTokensInput: document.getElementById("presetMaxTokensInput"),
  presetStreamInput: document.getElementById("presetStreamInput"),
  presetSystemPromptInput: document.getElementById("presetSystemPromptInput"),
  presetContextTemplateInput: document.getElementById("presetContextTemplateInput"),
  presetChatStartInput: document.getElementById("presetChatStartInput"),
  presetExampleSeparatorInput: document.getElementById("presetExampleSeparatorInput"),
  presetInstructEnabledInput: document.getElementById("presetInstructEnabledInput"),
  presetInstructUserInput: document.getElementById("presetInstructUserInput"),
  presetInstructAssistantInput: document.getElementById("presetInstructAssistantInput"),
  presetInstructSystemInput: document.getElementById("presetInstructSystemInput"),
  presetStopSequenceInput: document.getElementById("presetStopSequenceInput"),
  presetUseSystemPromptInput: document.getElementById("presetUseSystemPromptInput"),
  presetNamesBehaviorInput: document.getElementById("presetNamesBehaviorInput"),
  presetScenarioFormatInput: document.getElementById("presetScenarioFormatInput"),
  presetPersonalityFormatInput: document.getElementById("presetPersonalityFormatInput"),
  presetAddPromptBtn: document.getElementById("presetAddPromptBtn"),
  presetPromptList: document.getElementById("presetPromptList"),
  presetPromptModalBackdrop: document.getElementById("presetPromptModalBackdrop"),
  presetPromptEditor: document.getElementById("presetPromptEditor"),
  presetPromptIdentifierInput: document.getElementById("presetPromptIdentifierInput"),
  presetPromptNameInput: document.getElementById("presetPromptNameInput"),
  presetPromptRoleInput: document.getElementById("presetPromptRoleInput"),
  presetPromptInjectionPositionInput: document.getElementById("presetPromptInjectionPositionInput"),
  presetPromptInjectionDepthInput: document.getElementById("presetPromptInjectionDepthInput"),
  presetPromptEnabledInput: document.getElementById("presetPromptEnabledInput"),
  presetPromptSystemPromptInput: document.getElementById("presetPromptSystemPromptInput"),
  presetPromptMarkerInput: document.getElementById("presetPromptMarkerInput"),
  presetPromptForbidOverridesInput: document.getElementById("presetPromptForbidOverridesInput"),
  presetPromptContentInput: document.getElementById("presetPromptContentInput"),
  presetPromptSaveBtn: document.getElementById("presetPromptSaveBtn"),
  presetPromptDeleteBtn: document.getElementById("presetPromptDeleteBtn"),
  presetPromptCloseBtn: document.getElementById("presetPromptCloseBtn"),
  presetSaveBtn: document.getElementById("presetSaveBtn"),
  presetCancelBtn: document.getElementById("presetCancelBtn"),
  presetList: document.getElementById("presetList"),
  recentChats: document.getElementById("recentChats"),
  chatEmptyState: document.getElementById("chatEmptyState"),
  impersonateToggleBtn: document.getElementById("impersonateToggleBtn"),
  impersonateSelect: document.getElementById("impersonateSelect"),
  chatDrawerBackdrop: document.getElementById("chatDrawerBackdrop"),
  chatDrawer: document.getElementById("chatDrawer"),
  closeChatDrawerBtn: document.getElementById("closeChatDrawerBtn"),
  chatHamburgerBtn: document.getElementById("chatHamburgerBtn"),
  chatTitleInput: document.getElementById("chatTitleInput"),
  drawerPanels: Array.from(document.querySelectorAll("#chatDrawer .drawer-panel")),
  drawerBackBtns: Array.from(document.querySelectorAll("#chatDrawer .drawer-back-btn")),
  drawerHomeBtn: document.getElementById("drawerHomeBtn"),
  drawerProviderBtn: document.getElementById("drawerProviderBtn"),
  drawerPresetBtn: document.getElementById("drawerPresetBtn"),
  drawerMemoryBtn: document.getElementById("drawerMemoryBtn"),
  drawerLogBtn: document.getElementById("drawerLogBtn"),
  drawerMemoryLorebookBtn: document.getElementById("drawerMemoryLorebookBtn"),
  drawerMemorySummariseBtn: document.getElementById("drawerMemorySummariseBtn"),
  drawerLinks: Array.from(document.querySelectorAll("#chatDrawer .drawer-link[data-drawer-view]")),
  applyOtherPresetBtn: document.getElementById("applyOtherPresetBtn"),
  applyOtherProviderBtn: document.getElementById("applyOtherProviderBtn"),
  presetPicker: document.getElementById("presetPicker"),
  providerPicker: document.getElementById("providerPicker"),
  drawerPresetList: document.getElementById("drawerPresetList"),
  drawerProviderList: document.getElementById("drawerProviderList"),
  currentPresetLabel: document.getElementById("currentPresetLabel"),
  currentPresetSummaryBtn: document.getElementById("currentPresetSummaryBtn"),
  currentProviderLabel: document.getElementById("currentProviderLabel"),
  drawerPresetDetails: document.getElementById("drawerPresetDetails"),
  drawerProviderDetails: document.getElementById("drawerProviderDetails"),
  drawerPresetEditor: document.getElementById("drawerPresetEditor"),
  drawerPresetNameInput: document.getElementById("drawerPresetNameInput"),
  drawerPresetTempInput: document.getElementById("drawerPresetTempInput"),
  drawerPresetTopPInput: document.getElementById("drawerPresetTopPInput"),
  drawerPresetSystemInput: document.getElementById("drawerPresetSystemInput"),
  drawerPresetPromptList: document.getElementById("drawerPresetPromptList"),
  drawerPresetSaveBtn: document.getElementById("drawerPresetSaveBtn"),
  drawerPresetCancelBtn: document.getElementById("drawerPresetCancelBtn"),
  drawerProviderEditor: document.getElementById("drawerProviderEditor"),
  drawerProviderNameInput: document.getElementById("drawerProviderNameInput"),
  drawerProviderUrlInput: document.getElementById("drawerProviderUrlInput"),
  drawerProviderApiKeyInput: document.getElementById("drawerProviderApiKeyInput"),
  drawerProviderThinkStartInput: document.getElementById("drawerProviderThinkStartInput"),
  drawerProviderThinkEndInput: document.getElementById("drawerProviderThinkEndInput"),
  drawerProviderShowThinkingInput: document.getElementById("drawerProviderShowThinkingInput"),
  drawerProviderIncludeThinkingInput: document.getElementById("drawerProviderIncludeThinkingInput"),
  drawerProviderSaveBtn: document.getElementById("drawerProviderSaveBtn"),
  drawerProviderCancelBtn: document.getElementById("drawerProviderCancelBtn"),
  drawerLogList: document.getElementById("drawerLogList"),
  drawerLogTerminal: document.getElementById("drawerLogTerminal"),
  drawerLogMeta: document.getElementById("drawerLogMeta"),
  drawerLogCopyBtn: document.getElementById("drawerLogCopyBtn"),
  drawerLogClearBtn: document.getElementById("drawerLogClearBtn"),
  drawerLorebookImportBtn: document.getElementById("drawerLorebookImportBtn"),
  drawerLorebookSelect: document.getElementById("drawerLorebookSelect"),
  drawerLorebookEntries: document.getElementById("drawerLorebookEntries"),
  drawerLorebookSearchInput: document.getElementById("drawerLorebookSearchInput"),
  drawerLorebookList: document.getElementById("drawerLorebookList"),
  currentLorebookLabel: document.getElementById("currentLorebookLabel"),
  drawerLorebookScanDepthInput: document.getElementById("drawerLorebookScanDepthInput"),
  drawerLorebookTokenBudgetInput: document.getElementById("drawerLorebookTokenBudgetInput"),
  drawerLorebookRecursiveInput: document.getElementById("drawerLorebookRecursiveInput"),
  saveDrawerLorebookSettingsBtn: document.getElementById("saveDrawerLorebookSettingsBtn"),
  drawerLorebookStatus: document.getElementById("drawerLorebookStatus"),
  messageActionPopup: document.getElementById("messageActionPopup"),
  messageActionCard: document.getElementById("messageActionCard"),
  messageActionEditBtn: document.getElementById("messageActionEditBtn"),
  messageActionDeleteBtn: document.getElementById("messageActionDeleteBtn"),
  messageActionRegenerateBtn: document.getElementById("messageActionRegenerateBtn"),
  messageActionForkBtn: document.getElementById("messageActionForkBtn"),
  messageDeleteOptions: document.getElementById("messageDeleteOptions"),
  messageDeletePromptText: document.getElementById("messageDeletePromptText"),
  deleteOneMessageBtn: document.getElementById("deleteOneMessageBtn"),
  deleteFromHereBtn: document.getElementById("deleteFromHereBtn"),
  messageEditFieldWrap: document.getElementById("messageEditFieldWrap"),
  messageEditInput: document.getElementById("messageEditInput"),
  messageActionConfirmBtn: document.getElementById("messageActionConfirmBtn"),
  messageActionCloseBtn: document.getElementById("messageActionCloseBtn"),
  toastContainer: document.getElementById("toastContainer"),
  avatarSizeInput: document.getElementById("avatarSizeInput"),
  defaultTextColorInput: document.getElementById("defaultTextColorInput"),
  saveCustomizationBtn: document.getElementById("saveCustomizationBtn"),
  saveCustomizationTopBtn: document.getElementById("saveCustomizationTopBtn"),
  resetCustomizationBtn: document.getElementById("resetCustomizationBtn"),
  chatCustomizationSectionBtn: document.getElementById("chatCustomizationSectionBtn"),
  chatCustomizationSection: document.getElementById("chatCustomizationSection"),
  chatStyleSelect: document.getElementById("chatStyleSelect"),
  bubbleStyleSectionBtn: document.getElementById("bubbleStyleSectionBtn"),
  bubbleStyleSection: document.getElementById("bubbleStyleSection"),
  novelStyleSectionBtn: document.getElementById("novelStyleSectionBtn"),
  novelStyleSection: document.getElementById("novelStyleSection"),
  bubbleAssistantColorInput: document.getElementById("bubbleAssistantColorInput"),
  bubbleUserColorInput: document.getElementById("bubbleUserColorInput"),
  bubbleShowAssistantNameInput: document.getElementById("bubbleShowAssistantNameInput"),
  bubbleShowUserNameInput: document.getElementById("bubbleShowUserNameInput"),
  bubbleShowAssistantAvatarInput: document.getElementById("bubbleShowAssistantAvatarInput"),
  bubbleShowUserAvatarInput: document.getElementById("bubbleShowUserAvatarInput"),
  bubbleAssistantAvatarSizeInput: document.getElementById("bubbleAssistantAvatarSizeInput"),
  bubbleUserAvatarSizeInput: document.getElementById("bubbleUserAvatarSizeInput"),
  bubbleAssistantAvatarShapeInput: document.getElementById("bubbleAssistantAvatarShapeInput"),
  bubbleUserAvatarShapeInput: document.getElementById("bubbleUserAvatarShapeInput"),
  novelAssistantColorInput: document.getElementById("novelAssistantColorInput"),
  novelUserColorInput: document.getElementById("novelUserColorInput"),
  novelChatBackgroundColorInput: document.getElementById("novelChatBackgroundColorInput"),
  novelAssistantBlurInput: document.getElementById("novelAssistantBlurInput"),
  novelUserBlurInput: document.getElementById("novelUserBlurInput"),
  novelChatBackgroundBlurInput: document.getElementById("novelChatBackgroundBlurInput"),
  novelShowAssistantNameInput: document.getElementById("novelShowAssistantNameInput"),
  novelShowUserNameInput: document.getElementById("novelShowUserNameInput"),
  novelShowAssistantAvatarInput: document.getElementById("novelShowAssistantAvatarInput"),
  novelShowUserAvatarInput: document.getElementById("novelShowUserAvatarInput"),
  novelAssistantAvatarSizeInput: document.getElementById("novelAssistantAvatarSizeInput"),
  novelUserAvatarSizeInput: document.getElementById("novelUserAvatarSizeInput"),
  novelAssistantAvatarShapeInput: document.getElementById("novelAssistantAvatarShapeInput"),
  novelUserAvatarShapeInput: document.getElementById("novelUserAvatarShapeInput"),
  novelAssistantAvatarAlignInput: document.getElementById("novelAssistantAvatarAlignInput"),
  novelUserAvatarAlignInput: document.getElementById("novelUserAvatarAlignInput"),
  novelAssistantTextFlowInput: document.getElementById("novelAssistantTextFlowInput"),
  novelUserTextFlowInput: document.getElementById("novelUserTextFlowInput"),
  novelAssistantImageBtn: document.getElementById("novelAssistantImageBtn"),
  novelUserImageBtn: document.getElementById("novelUserImageBtn"),
  novelChatImageBtn: document.getElementById("novelChatImageBtn"),
  novelAssistantImageInput: document.getElementById("novelAssistantImageInput"),
  novelUserImageInput: document.getElementById("novelUserImageInput"),
  novelChatImageInput: document.getElementById("novelChatImageInput"),
  novelAssistantImageMeta: document.getElementById("novelAssistantImageMeta"),
  novelUserImageMeta: document.getElementById("novelUserImageMeta"),
  novelChatImageMeta: document.getElementById("novelChatImageMeta"),
  avatarSizeSectionBtn: document.getElementById("avatarSizeSectionBtn"),
  avatarSizeSection: document.getElementById("avatarSizeSection"),
  textWrappingSectionBtn: document.getElementById("textWrappingSectionBtn"),
  textWrappingSection: document.getElementById("textWrappingSection"),
  wrapRuleSearchInput: document.getElementById("wrapRuleSearchInput"),
  wrapRuleList: document.getElementById("wrapRuleList"),
  wrapRuleEditor: document.getElementById("wrapRuleEditor"),
  wrapRuleTitleInput: document.getElementById("wrapRuleTitleInput"),
  wrapRuleStartInput: document.getElementById("wrapRuleStartInput"),
  wrapRuleEndInput: document.getElementById("wrapRuleEndInput"),
  wrapRuleColorInput: document.getElementById("wrapRuleColorInput"),
  wrapRuleHexInput: document.getElementById("wrapRuleHexInput"),
  addWrapRuleBtn: document.getElementById("addWrapRuleBtn"),
  saveWrapRuleBtn: document.getElementById("saveWrapRuleBtn"),
  deleteWrapRuleBtn: document.getElementById("deleteWrapRuleBtn"),
  closeWrapRuleEditorBtn: document.getElementById("closeWrapRuleEditorBtn"),
  textFontSectionBtn: document.getElementById("textFontSectionBtn"),
  textFontSection: document.getElementById("textFontSection"),
  openFontManagerBtn: document.getElementById("openFontManagerBtn"),
  currentFontName: document.getElementById("currentFontName"),
  fontManagerBackdrop: document.getElementById("fontManagerBackdrop"),
  fontManagerCard: document.getElementById("fontManagerCard"),
  closeFontManagerBtn: document.getElementById("closeFontManagerBtn"),
  fontFileInput: document.getElementById("fontFileInput"),
  browseFontFileBtn: document.getElementById("browseFontFileBtn"),
  uploadedFontMeta: document.getElementById("uploadedFontMeta"),
  googleFontSearchInput: document.getElementById("googleFontSearchInput"),
  addGoogleFontBtn: document.getElementById("addGoogleFontBtn"),
  fontManagerStatus: document.getElementById("fontManagerStatus"),
  resetAppFontBtn: document.getElementById("resetAppFontBtn"),
  exportDataBtn: document.getElementById("exportDataBtn"),
  importDataBtn: document.getElementById("importDataBtn"),
  importDataInput: document.getElementById("importDataInput"),
  dataStatus: document.getElementById("dataStatus"),

  // Lorebook V2 nodes
  lbv2Panel: document.querySelector("#libraryLorebooksPanel.lorebook-v2-panel"),
  lbv2Stack: document.getElementById("lbv2Stack"),
  lbv2Level1: document.getElementById("lbv2Level1"),
  lbv2Level2: document.getElementById("lbv2Level2"),
  lbv2Level3: document.getElementById("lbv2Level3"),
  lbv2BackLibraryBtn: document.getElementById("lbv2BackLibraryBtn"),
  lbv2CreateLorebookTopBtn: document.getElementById("lbv2CreateLorebookTopBtn"),
  lbv2LorebookOverflowBtn: document.getElementById("lbv2LorebookOverflowBtn"),
  lbv2LorebookSearchInput: document.getElementById("lbv2LorebookSearchInput"),
  lbv2LorebookList: document.getElementById("lbv2LorebookList"),
  lbv2LorebookEmptyState: document.getElementById("lbv2LorebookEmptyState"),
  lbv2EmptyCreateBtn: document.getElementById("lbv2EmptyCreateBtn"),
  lbv2LorebookFab: document.getElementById("lbv2LorebookFab"),
  lbv2EntriesTopbar: document.getElementById("lbv2EntriesTopbar"),
  lbv2BulkTopbar: document.getElementById("lbv2BulkTopbar"),
  lbv2EntriesBackBtn: document.getElementById("lbv2EntriesBackBtn"),
  lbv2EntriesBackLabel: document.getElementById("lbv2EntriesBackLabel"),
  lbv2EntriesTitle: document.getElementById("lbv2EntriesTitle"),
  lbv2LorebookSettingsBtn: document.getElementById("lbv2LorebookSettingsBtn"),
  lbv2EntriesOverflowBtn: document.getElementById("lbv2EntriesOverflowBtn"),
  lbv2BulkCancelBtn: document.getElementById("lbv2BulkCancelBtn"),
  lbv2BulkTitle: document.getElementById("lbv2BulkTitle"),
  lbv2BulkActionsBtn: document.getElementById("lbv2BulkActionsBtn"),
  lbv2EntrySearchInput: document.getElementById("lbv2EntrySearchInput"),
  lbv2FilterChipRow: document.getElementById("lbv2FilterChipRow"),
  lbv2EntryList: document.getElementById("lbv2EntryList"),
  lbv2EntryEmptyState: document.getElementById("lbv2EntryEmptyState"),
  lbv2EmptyEntryCreateBtn: document.getElementById("lbv2EmptyEntryCreateBtn"),
  lbv2EntryFooter: document.getElementById("lbv2EntryFooter"),
  lbv2EntryCountLabel: document.getElementById("lbv2EntryCountLabel"),
  lbv2EntryTokenLabel: document.getElementById("lbv2EntryTokenLabel"),
  lbv2EntryNewBtn: document.getElementById("lbv2EntryNewBtn"),
  lbv2EditorBackBtn: document.getElementById("lbv2EditorBackBtn"),
  lbv2EditorTopTitle: document.getElementById("lbv2EditorTopTitle"),
  lbv2SaveIndicator: document.getElementById("lbv2SaveIndicator"),
  lbv2EditorOverflowBtn: document.getElementById("lbv2EditorOverflowBtn"),
  lbv2EditorBody: document.getElementById("lbv2EditorBody"),
  lbv2Accordions: Array.from(document.querySelectorAll("#lbv2EditorBody .lbv2-accordion")),
  lbv2EntryNameInput: document.getElementById("lbv2EntryNameInput"),
  lbv2EntryCommentInput: document.getElementById("lbv2EntryCommentInput"),
  lbv2PrimaryKeysInput: document.getElementById("lbv2PrimaryKeysInput"),
  lbv2PrimaryKeysPreview: document.getElementById("lbv2PrimaryKeysPreview"),
  lbv2SecondaryKeysInput: document.getElementById("lbv2SecondaryKeysInput"),
  lbv2ActivationAnyInput: document.getElementById("lbv2ActivationAnyInput"),
  lbv2ActivationAllInput: document.getElementById("lbv2ActivationAllInput"),
  lbv2ActivationAndOrInput: document.getElementById("lbv2ActivationAndOrInput"),
  lbv2ActivationNotInput: document.getElementById("lbv2ActivationNotInput"),
  lbv2CaseSensitiveInput: document.getElementById("lbv2CaseSensitiveInput"),
  lbv2WholeWordsInput: document.getElementById("lbv2WholeWordsInput"),
  lbv2RegexInput: document.getElementById("lbv2RegexInput"),
  lbv2RegexHelper: document.getElementById("lbv2RegexHelper"),
  lbv2ScanDepthInput: document.getElementById("lbv2ScanDepthInput"),
  lbv2PerEntryTokenBudgetInput: document.getElementById("lbv2PerEntryTokenBudgetInput"),
  lbv2EntryContentInput: document.getElementById("lbv2EntryContentInput"),
  lbv2ContentTokenCount: document.getElementById("lbv2ContentTokenCount"),
  lbv2TriggerWarning: document.getElementById("lbv2TriggerWarning"),
  lbv2InsertionPositionBtn: document.getElementById("lbv2InsertionPositionBtn"),
  lbv2InsertionPositionInput: document.getElementById("lbv2InsertionPositionInput"),
  lbv2InsertionDepthField: document.getElementById("lbv2InsertionDepthField"),
  lbv2InsertionDepthInput: document.getElementById("lbv2InsertionDepthInput"),
  lbv2PriorityInput: document.getElementById("lbv2PriorityInput"),
  lbv2EnabledInput: document.getElementById("lbv2EnabledInput"),
  lbv2ConstantHelper: document.getElementById("lbv2ConstantHelper"),
  lbv2PreventRecursionInput: document.getElementById("lbv2PreventRecursionInput"),
  lbv2ExcludeFromRecursionInput: document.getElementById("lbv2ExcludeFromRecursionInput"),
  lbv2ActivationProbabilitySlider: document.getElementById("lbv2ActivationProbabilitySlider"),
  lbv2ActivationProbabilityInput: document.getElementById("lbv2ActivationProbabilityInput"),
  lbv2GroupInput: document.getElementById("lbv2GroupInput"),
  lbv2GroupWeightField: document.getElementById("lbv2GroupWeightField"),
  lbv2GroupWeightInput: document.getElementById("lbv2GroupWeightInput"),
  lbv2CategoryPickerBtn: document.getElementById("lbv2CategoryPickerBtn"),
  lbv2CategoryInput: document.getElementById("lbv2CategoryInput"),
  lbv2TagList: document.getElementById("lbv2TagList"),
  lbv2TagInput: document.getElementById("lbv2TagInput"),
  lbv2TagSuggestionList: document.getElementById("lbv2TagSuggestionList"),
  lbv2AutomationIdInput: document.getElementById("lbv2AutomationIdInput"),
  lbv2LinkedEntryList: document.getElementById("lbv2LinkedEntryList"),
  lbv2LinkEntryBtn: document.getElementById("lbv2LinkEntryBtn"),
  lbv2StickyInput: document.getElementById("lbv2StickyInput"),
  lbv2CooldownInput: document.getElementById("lbv2CooldownInput"),
  lbv2DelayInput: document.getElementById("lbv2DelayInput"),
  lbv2TurnStartInput: document.getElementById("lbv2TurnStartInput"),
  lbv2TurnEndInput: document.getElementById("lbv2TurnEndInput"),
  lbv2EntryMetadata: document.getElementById("lbv2EntryMetadata"),

  lbv2SettingsModal: document.getElementById("lbv2SettingsModal"),
  lbv2SettingsBackBtn: document.getElementById("lbv2SettingsBackBtn"),
  lbv2SettingsNameInput: document.getElementById("lbv2SettingsNameInput"),
  lbv2SettingsDescriptionInput: document.getElementById("lbv2SettingsDescriptionInput"),
  lbv2SettingsColorDots: document.getElementById("lbv2SettingsColorDots"),
  lbv2SettingsTokenBudgetInput: document.getElementById("lbv2SettingsTokenBudgetInput"),
  lbv2SettingsDefaultPositionBtn: document.getElementById("lbv2SettingsDefaultPositionBtn"),
  lbv2SettingsDefaultPositionInput: document.getElementById("lbv2SettingsDefaultPositionInput"),
  lbv2SettingsDefaultScanDepthInput: document.getElementById("lbv2SettingsDefaultScanDepthInput"),
  lbv2SettingsDefaultOrderInput: document.getElementById("lbv2SettingsDefaultOrderInput"),
  lbv2SettingsDefaultEnabledInput: document.getElementById("lbv2SettingsDefaultEnabledInput"),
  lbv2SettingsDefaultCaseInput: document.getElementById("lbv2SettingsDefaultCaseInput"),
  lbv2SettingsRecursiveInput: document.getElementById("lbv2SettingsRecursiveInput"),
  lbv2SettingsRecursionDepthField: document.getElementById("lbv2SettingsRecursionDepthField"),
  lbv2SettingsRecursionDepthInput: document.getElementById("lbv2SettingsRecursionDepthInput"),
  lbv2AttachmentCharacterField: document.getElementById("lbv2AttachmentCharacterField"),
  lbv2AttachmentCharacterSelect: document.getElementById("lbv2AttachmentCharacterSelect"),
  lbv2AttachmentChatWrap: document.getElementById("lbv2AttachmentChatWrap"),
  lbv2AttachmentChatChipRow: document.getElementById("lbv2AttachmentChatChipRow"),
  lbv2AttachChatBtn: document.getElementById("lbv2AttachChatBtn"),
  lbv2ExportJsonBtn: document.getElementById("lbv2ExportJsonBtn"),
  lbv2ExportStBtn: document.getElementById("lbv2ExportStBtn"),
  lbv2ImportEntriesBtn: document.getElementById("lbv2ImportEntriesBtn"),
  lbv2DisableAllEntriesBtn: document.getElementById("lbv2DisableAllEntriesBtn"),
  lbv2EnableAllEntriesBtn: document.getElementById("lbv2EnableAllEntriesBtn"),
  lbv2DeleteLorebookBtn: document.getElementById("lbv2DeleteLorebookBtn"),

  lbv2SheetBackdrop: document.getElementById("lbv2SheetBackdrop"),
  lbv2Sheet: document.getElementById("lbv2Sheet"),
  lbv2SheetTitle: document.getElementById("lbv2SheetTitle"),
  lbv2SheetContent: document.getElementById("lbv2SheetContent"),
  lbv2DialogBackdrop: document.getElementById("lbv2DialogBackdrop"),
  lbv2DialogTitle: document.getElementById("lbv2DialogTitle"),
  lbv2DialogBody: document.getElementById("lbv2DialogBody"),
  lbv2DialogCancelBtn: document.getElementById("lbv2DialogCancelBtn"),
  lbv2DialogConfirmBtn: document.getElementById("lbv2DialogConfirmBtn"),
  lbv2EntryImportInput: document.getElementById("lbv2EntryImportInput"),
};

void init();

async function init() {
  state = await loadState();
  ensureCollections();
  syncInputsFromState();
  renderModelOptions();
  applyCustomization();
  renderMessages();
  renderCharacters();
  renderProviders();
  renderPresets();
  renderLorebooks();
  setActiveLibraryPanel("libraryHubPanel");
  renderLibraryHub();
  initLorebookV2();
  renderRecentChats();
  renderImpersonationOptions();
  renderDrawerState();
  syncCustomizationInputs();
  syncLorebookSettingsInputs();
  renderWrapRuleList();
  applyWelcomeState();
  updateHomeGreeting();
  switchView(state.activeView || "homeView");
  bindEvents();
  autoResizeTextarea(els.messageInput);
  updateChatUiVisibility();
  maybeAutoFetchModels();
  persistState();
}

function ensureCollections() {
  state.characters = Array.isArray(state.characters)
    ? state.characters.map((character) => normalizeCharacterRecord(character))
    : [];
  state.providers = Array.isArray(state.providers) ? state.providers : [];
  state.assets = Array.isArray(state.assets) ? state.assets : [];
  state.lorebooks = Array.isArray(state.lorebooks)
    ? state.lorebooks.map((lorebook) => normalizeLorebookRecord(lorebook))
    : [];
  state.lorebookSettings = {
    ...structuredClone(defaults.lorebookSettings),
    ...(state.lorebookSettings && typeof state.lorebookSettings === "object" ? state.lorebookSettings : {}),
  };
  state.lorebookSettings.scanDepth = Math.min(64, Math.max(1, Number.parseInt(state.lorebookSettings.scanDepth, 10) || defaults.lorebookSettings.scanDepth));
  state.lorebookSettings.tokenBudget = Math.min(8000, Math.max(64, Number.parseInt(state.lorebookSettings.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget));
  state.lorebookSettings.recursiveScanning = state.lorebookSettings.recursiveScanning !== false;
  state.presets = Array.isArray(state.presets) && state.presets.length ? state.presets : structuredClone(defaults.presets);
  state.presets = state.presets.map(normalizePreset);
  state.chats = Array.isArray(state.chats)
    ? state.chats.map((chat) => normalizeChatRecord(chat))
    : [];
  state.recentChats = Array.isArray(state.recentChats) ? state.recentChats : [];
  state.currentChatId = typeof state.currentChatId === "string" ? state.currentChatId : "";
  state.currentChatTitle = typeof state.currentChatTitle === "string" && state.currentChatTitle.trim() ? state.currentChatTitle : "New chat";
  state.lastGenerationType = typeof state.lastGenerationType === "string" ? state.lastGenerationType : "";
  state.promptLogs = Array.isArray(state.promptLogs)
    ? state.promptLogs.filter((entry) => entry && typeof entry === "object")
      .slice(0, 5)
      .map((entry) => ({
        id: typeof entry.id === "string" ? entry.id : crypto.randomUUID(),
        createdAt: typeof entry.createdAt === "string" ? entry.createdAt : new Date().toISOString(),
        model: typeof entry.model === "string" ? entry.model : "",
        endpoint: typeof entry.endpoint === "string" ? entry.endpoint : "",
        chatId: typeof entry.chatId === "string" ? entry.chatId : "",
        chatTitle: typeof entry.chatTitle === "string" ? entry.chatTitle : "",
        payload: entry.payload && typeof entry.payload === "object" ? JSON.parse(JSON.stringify(entry.payload)) : {},
      }))
    : [];
  state.customization = {
    ...structuredClone(defaults.customization),
    ...(state.customization && typeof state.customization === "object" ? state.customization : {}),
  };
  state.macroGlobalVars = state.macroGlobalVars && typeof state.macroGlobalVars === "object" ? state.macroGlobalVars : {};
  state.macroChatVars = state.macroChatVars && typeof state.macroChatVars === "object" ? state.macroChatVars : {};
  state.customization.wrapRules = Array.isArray(state.customization.wrapRules) && state.customization.wrapRules.length
    ? state.customization.wrapRules
    : structuredClone(defaults.customization.wrapRules);
  state.customization.wrapRules = state.customization.wrapRules.map((rule, index) => ({
    id: typeof rule?.id === "string" && rule.id.trim() ? rule.id : crypto.randomUUID(),
    title: typeof rule?.title === "string" ? rule.title : `Wrap ${index + 1}`,
    start: String(rule?.start ?? ""),
    end: String(rule?.end ?? ""),
    color: isValidHexColor(rule?.color) ? rule.color : "#c084fc",
    enabled: rule?.enabled !== false,
  })).filter((rule) => rule.start && rule.end);
  if (!state.customization.wrapRules.length) {
    state.customization.wrapRules = structuredClone(defaults.customization.wrapRules);
  }
  state.customization.appFontSource = typeof state.customization.appFontSource === "string" ? state.customization.appFontSource : defaults.customization.appFontSource;
  if (!["system", "file", "google"].includes(state.customization.appFontSource)) {
    state.customization.appFontSource = defaults.customization.appFontSource;
  }
  state.customization.appFontFamily = typeof state.customization.appFontFamily === "string" && state.customization.appFontFamily.trim()
    ? state.customization.appFontFamily.trim()
    : defaults.customization.appFontFamily;
  state.customization.appFontFileName = typeof state.customization.appFontFileName === "string" ? state.customization.appFontFileName : "";
  state.customization.appFontDataUrl = typeof state.customization.appFontDataUrl === "string" ? state.customization.appFontDataUrl : "";
  state.customization.appFontGoogleName = typeof state.customization.appFontGoogleName === "string" ? state.customization.appFontGoogleName : "";
  state.customization.appFontGoogleCss = typeof state.customization.appFontGoogleCss === "string" ? state.customization.appFontGoogleCss : "";

  state.customization.chatStyle = state.customization.chatStyle === "novel" ? "novel" : "bubble";

  const bubbleSource = state.customization.bubble && typeof state.customization.bubble === "object"
    ? state.customization.bubble
    : {};
  const novelSource = state.customization.novel && typeof state.customization.novel === "object"
    ? state.customization.novel
    : {};

  state.customization.bubble = {
    ...structuredClone(defaults.customization.bubble),
    ...bubbleSource,
  };
  state.customization.novel = {
    ...structuredClone(defaults.customization.novel),
    ...novelSource,
  };

  const sanitizeAvatarShape = (value, fallback) => (["cutout", "circle", "square", "rectangle"].includes(value) ? value : fallback);
  const sanitizeAvatarAlign = (value, fallback) => (["left", "center", "right"].includes(value) ? value : fallback);
  const sanitizeTextFlow = (value, fallback) => (["wrap", "below"].includes(value) ? value : fallback);

  state.customization.bubble.assistantBubbleColor = isValidHexColor(state.customization.bubble.assistantBubbleColor)
    ? state.customization.bubble.assistantBubbleColor
    : defaults.customization.bubble.assistantBubbleColor;
  state.customization.bubble.userBubbleColor = isValidHexColor(state.customization.bubble.userBubbleColor)
    ? state.customization.bubble.userBubbleColor
    : defaults.customization.bubble.userBubbleColor;
  state.customization.bubble.showAssistantName = state.customization.bubble.showAssistantName !== false;
  state.customization.bubble.showUserName = state.customization.bubble.showUserName !== false;
  state.customization.bubble.showAssistantAvatar = state.customization.bubble.showAssistantAvatar !== false;
  state.customization.bubble.showUserAvatar = state.customization.bubble.showUserAvatar !== false;
  state.customization.bubble.assistantAvatarSize = Math.min(160, Math.max(24, Number.parseInt(state.customization.bubble.assistantAvatarSize, 10) || state.customization.avatarSize || defaults.customization.bubble.assistantAvatarSize));
  state.customization.bubble.userAvatarSize = Math.min(160, Math.max(24, Number.parseInt(state.customization.bubble.userAvatarSize, 10) || state.customization.avatarSize || defaults.customization.bubble.userAvatarSize));
  state.customization.bubble.assistantAvatarShape = sanitizeAvatarShape(state.customization.bubble.assistantAvatarShape, defaults.customization.bubble.assistantAvatarShape);
  state.customization.bubble.userAvatarShape = sanitizeAvatarShape(state.customization.bubble.userAvatarShape, defaults.customization.bubble.userAvatarShape);

  state.customization.novel.showAssistantName = state.customization.novel.showAssistantName !== false;
  state.customization.novel.showUserName = state.customization.novel.showUserName !== false;
  state.customization.novel.showAssistantAvatar = state.customization.novel.showAssistantAvatar !== false;
  state.customization.novel.showUserAvatar = state.customization.novel.showUserAvatar !== false;
  state.customization.novel.assistantAvatarSize = Math.min(200, Math.max(24, Number.parseInt(state.customization.novel.assistantAvatarSize, 10) || defaults.customization.novel.assistantAvatarSize));
  state.customization.novel.userAvatarSize = Math.min(200, Math.max(24, Number.parseInt(state.customization.novel.userAvatarSize, 10) || defaults.customization.novel.userAvatarSize));
  state.customization.novel.assistantAvatarShape = sanitizeAvatarShape(state.customization.novel.assistantAvatarShape, defaults.customization.novel.assistantAvatarShape);
  state.customization.novel.userAvatarShape = sanitizeAvatarShape(state.customization.novel.userAvatarShape, defaults.customization.novel.userAvatarShape);
  state.customization.novel.assistantAvatarAlign = sanitizeAvatarAlign(state.customization.novel.assistantAvatarAlign, defaults.customization.novel.assistantAvatarAlign);
  state.customization.novel.userAvatarAlign = sanitizeAvatarAlign(state.customization.novel.userAvatarAlign, defaults.customization.novel.userAvatarAlign);
  state.customization.novel.assistantTextFlow = sanitizeTextFlow(state.customization.novel.assistantTextFlow, defaults.customization.novel.assistantTextFlow);
  state.customization.novel.userTextFlow = sanitizeTextFlow(state.customization.novel.userTextFlow, defaults.customization.novel.userTextFlow);
  state.customization.novel.assistantBackgroundColor = isValidHexColor(state.customization.novel.assistantBackgroundColor)
    ? state.customization.novel.assistantBackgroundColor
    : defaults.customization.novel.assistantBackgroundColor;
  state.customization.novel.userBackgroundColor = isValidHexColor(state.customization.novel.userBackgroundColor)
    ? state.customization.novel.userBackgroundColor
    : defaults.customization.novel.userBackgroundColor;
  state.customization.novel.chatBackgroundColor = isValidHexColor(state.customization.novel.chatBackgroundColor)
    ? state.customization.novel.chatBackgroundColor
    : defaults.customization.novel.chatBackgroundColor;
  const normalizedNovelChatBg = String(state.customization.novel.chatBackgroundColor || "").trim().toLowerCase();
  if (["#0f0e11", "#100f14", "#0d0d0f"].includes(normalizedNovelChatBg)) {
    state.customization.novel.chatBackgroundColor = defaults.customization.novel.chatBackgroundColor;
  }
  state.customization.novel.assistantBackgroundImage = typeof state.customization.novel.assistantBackgroundImage === "string" ? state.customization.novel.assistantBackgroundImage : "";
  state.customization.novel.userBackgroundImage = typeof state.customization.novel.userBackgroundImage === "string" ? state.customization.novel.userBackgroundImage : "";
  state.customization.novel.chatBackgroundImage = typeof state.customization.novel.chatBackgroundImage === "string" ? state.customization.novel.chatBackgroundImage : "";
  state.customization.novel.assistantBackgroundBlur = Math.min(32, Math.max(0, Number.parseInt(state.customization.novel.assistantBackgroundBlur, 10) || 0));
  state.customization.novel.userBackgroundBlur = Math.min(32, Math.max(0, Number.parseInt(state.customization.novel.userBackgroundBlur, 10) || 0));
  state.customization.novel.chatBackgroundBlur = Math.min(32, Math.max(0, Number.parseInt(state.customization.novel.chatBackgroundBlur, 10) || 0));

  if (state.activeProviderId && !state.providers.some((provider) => provider.id === state.activeProviderId)) {
    state.activeProviderId = "";
  }

  if ((!state.endpoint || !state.apiKey) && state.activeProviderId) {
    const activeProvider = state.providers.find((provider) => provider.id === state.activeProviderId);
    if (activeProvider) {
      if (!state.endpoint) state.endpoint = activeProvider.baseUrl || "";
      if (!state.apiKey) state.apiKey = activeProvider.apiKey || "";
    }
  }

  migrateLegacyChats();
}

function bindEvents() {
  els.enterAppBtn?.addEventListener("click", enterApp);
  els.homeBrandBtn?.addEventListener("click", goHome);
  els.settingsToggle?.addEventListener("click", () => {
    if (state.activeView === "chatsView") {
      toggleChatDrawer(true);
    }
  });
  els.closeSettingsBtn?.addEventListener("click", () => toggleSettings(false));
  els.saveSettingsBtn.addEventListener("click", saveSettings);
  els.clearStorageBtn.addEventListener("click", clearStorage);
  els.fetchModelsBtn?.addEventListener("click", () => fetchModelsIntoSettings());
  els.modelSelect.addEventListener("change", () => {
    state.selectedModel = els.modelSelect.value;
    persistState();
  });
  els.composer.addEventListener("submit", (event) => event.preventDefault());
  els.sendBtn.addEventListener("click", onSend);
  els.stopBtn.addEventListener("click", stopRequest);
  els.newChatBtn.addEventListener("click", openLastInteractedChat);
  els.messageInput.addEventListener("input", () => autoResizeTextarea(els.messageInput));
  els.messageInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
    }
  });
  els.navCards.forEach((card) => {
    card.addEventListener("click", () => {
      switchView(card.dataset.view);
      if (card.dataset.view === "libraryView") {
        setActiveLibraryPanel("libraryHubPanel");
        renderLibraryHub();
      }
      if (card.dataset.view === "chatsView") {
        state.currentChatId = "";
        state.messages = [];
        state.currentChatTitle = "New chat";
        refreshRecentChatsFromState();
        renderMessages();
        renderRecentChats();
        renderDrawerState();
        updateChatUiVisibility();
      }
    });
  });
  els.characterSearchInput.addEventListener("input", renderCharacters);
  els.characterStudioBtn.addEventListener("click", createCharacter);
  els.characterStudioImportJsonBtn?.addEventListener("click", () => els.characterImportInput.click());
  els.characterStudioImportAvatarBtn?.addEventListener("click", () => els.characterAvatarImportInput.click());
  els.characterSaveBtn.addEventListener("click", saveCharacterFromEditor);
  els.characterCancelBtn.addEventListener("click", hideCharacterEditor);
  els.characterImportInput.addEventListener("change", importCharacters);
  els.characterAvatarImportInput.addEventListener("change", importCharacterFromAvatar);
  els.characterAvatarInput.addEventListener("change", onCharacterAvatarSelected);
  els.libraryCategoryChips?.forEach((chip) => {
    chip.addEventListener("click", () => setActiveLibraryPanel(chip.dataset.libraryPanel || "libraryHubPanel"));
  });

  els.libraryPanelBackBtns?.forEach((button) => {
    button.addEventListener("click", () => setActiveLibraryPanel("libraryHubPanel"));
  });

  els.libraryHubSearchInput?.addEventListener("input", () => {
    libraryHubSearchQuery = String(els.libraryHubSearchInput.value || "");
    renderLibraryHub();
  });
  bindLorebookV2Events();
  els.lorebookSearchInput?.addEventListener("input", renderLorebooks);
  els.lorebookImportBtn?.addEventListener("click", () => els.lorebookImportInput?.click());
  els.lorebookImportInput?.addEventListener("change", importLorebooks);
  els.lorebookExportBtn?.addEventListener("click", exportLorebooks);
  els.lorebookNewBtn?.addEventListener("click", () => showLorebookForm());
  els.lorebookEditorImportBtn?.addEventListener("click", () => els.lorebookEditorImportInput?.click());
  els.lorebookEditorImportInput?.addEventListener("change", importLorebookIntoEditor);
  els.lorebookEditorBackBtn?.addEventListener("click", hideLorebookForm);
  els.lorebookSaveBtn?.addEventListener("click", saveLorebookFromForm);
  els.lorebookCancelBtn?.addEventListener("click", hideLorebookForm);
  els.lorebookEntrySearchInput?.addEventListener("input", () => {
    lorebookEntrySearchQuery = els.lorebookEntrySearchInput.value || "";
    renderLorebookEntryDraft();
  });
  els.lorebookEntryNewBtn?.addEventListener("click", () => openLorebookEntryEditor());
  els.lorebookEntrySaveBtn?.addEventListener("click", saveLorebookEntryFromEditor);
  els.lorebookEntryDeleteBtn?.addEventListener("click", deleteLorebookEntryFromEditor);
  els.lorebookEntryCancelBtn?.addEventListener("click", closeLorebookEntryEditor);
  els.saveLorebookSettingsBtn?.addEventListener("click", saveLorebookSettingsFromInputs);
  els.drawerLorebookSearchInput?.addEventListener("input", renderDrawerLorebooks);
  els.saveDrawerLorebookSettingsBtn?.addEventListener("click", saveDrawerLorebookSettings);
  els.providerSearchInput.addEventListener("input", renderProviders);
  els.providerNewBtn.addEventListener("click", showProviderForm);
  els.providerCancelBtn.addEventListener("click", hideProviderForm);
  els.providerSaveBtn.addEventListener("click", saveProviderFromForm);
  els.providerFetchModelsBtn?.addEventListener("click", () => fetchModelsIntoProviderForm());
  els.providerImportBtn.addEventListener("click", () => els.providerImportInput.click());
  els.providerImportInput.addEventListener("change", importProviders);
  els.providerExportBtn.addEventListener("click", exportProviders);
  els.presetImportBtn?.addEventListener("click", () => els.presetImportInput.click());
  els.presetImportInput?.addEventListener("change", importPresets);
  els.presetExportBtn?.addEventListener("click", exportPresets);
  els.presetNewBtn?.addEventListener("click", () => showPresetForm());
  els.presetCancelBtn?.addEventListener("click", hidePresetForm);
  els.presetSaveBtn?.addEventListener("click", savePresetFromForm);
  els.presetAddPromptBtn?.addEventListener("click", createPresetPromptDraft);
  els.presetPromptSaveBtn?.addEventListener("click", savePresetPromptDraft);
  els.presetPromptDeleteBtn?.addEventListener("click", deletePresetPromptDraft);
  els.presetPromptCloseBtn?.addEventListener("click", hidePresetPromptEditor);
  els.presetPromptModalBackdrop?.addEventListener("click", hidePresetPromptEditor);
  els.closeChatDrawerBtn?.addEventListener("click", () => toggleChatDrawer(false));
  els.chatHamburgerBtn?.addEventListener("click", () => toggleChatDrawer(true));
  els.chatDrawerBackdrop?.addEventListener("click", () => toggleChatDrawer(false));
  els.drawerLinks.forEach((button) => {
    button.addEventListener("click", () => {
      if (!button.dataset.drawerView) return;
      switchView(button.dataset.drawerView);
      toggleChatDrawer(false);
    });
  });
  els.drawerHomeBtn?.addEventListener("click", () => {
    switchView("homeView");
    toggleChatDrawer(false);
  });
  els.drawerProviderBtn?.addEventListener("click", () => openDrawerPanel("provider"));
  els.drawerPresetBtn?.addEventListener("click", () => openDrawerPanel("preset"));
  els.drawerMemoryBtn?.addEventListener("click", () => openDrawerPanel("memory"));
  els.drawerLogBtn?.addEventListener("click", () => openDrawerPanel("log"));
  els.drawerMemoryLorebookBtn?.addEventListener("click", () => openDrawerPanel("lorebook"));
  els.drawerMemorySummariseBtn?.addEventListener("click", () => showToast("Summarise coming soon", "success"));
  els.drawerBackBtns?.forEach((button) => button.addEventListener("click", closeDrawerPanel));
  els.drawerLorebookImportBtn?.addEventListener("click", () => els.lorebookImportInput?.click());
  els.drawerLorebookSelect?.addEventListener("change", () => {
    drawerLorebookSelectedId = els.drawerLorebookSelect?.value || "";
    renderDrawerLorebooks();
  });
  els.currentPresetSummaryBtn?.addEventListener("click", () => {
    openDrawerPanel("preset-editor");
  });
  els.chatTitleInput?.addEventListener("input", onChatTitleInput);
  els.impersonateToggleBtn?.addEventListener("click", toggleImpersonationPicker);
  els.impersonateSelect?.addEventListener("change", onImpersonationChange);
  els.saveCustomizationBtn?.addEventListener("click", saveCustomization);
  els.saveCustomizationTopBtn?.addEventListener("click", saveCustomization);
  els.resetCustomizationBtn?.addEventListener("click", resetCustomization);
  els.chatCustomizationSectionBtn?.addEventListener("click", () => toggleSection(els.chatCustomizationSection));
  els.chatStyleSelect?.addEventListener("change", onChatStyleSelectChange);
  els.bubbleStyleSectionBtn?.addEventListener("click", () => toggleSection(els.bubbleStyleSection));
  els.novelStyleSectionBtn?.addEventListener("click", () => toggleSection(els.novelStyleSection));
  els.avatarSizeSectionBtn?.addEventListener("click", () => toggleSection(els.avatarSizeSection));
  els.textWrappingSectionBtn?.addEventListener("click", () => toggleSection(els.textWrappingSection));
  els.addWrapRuleBtn?.addEventListener("click", createWrapRule);
  els.saveWrapRuleBtn?.addEventListener("click", saveWrapRuleFromEditor);
  els.deleteWrapRuleBtn?.addEventListener("click", deleteWrapRuleFromEditor);
  els.closeWrapRuleEditorBtn?.addEventListener("click", hideWrapRuleEditor);
  els.textFontSectionBtn?.addEventListener("click", () => toggleSection(els.textFontSection));
  els.openFontManagerBtn?.addEventListener("click", openTextFontManager);
  els.closeFontManagerBtn?.addEventListener("click", () => toggleFontManager(false));
  els.fontManagerBackdrop?.addEventListener("click", (event) => {
    if (event.target === els.fontManagerBackdrop) {
      toggleFontManager(false);
    }
  });
  els.browseFontFileBtn?.addEventListener("click", () => els.fontFileInput?.click());
  els.fontFileInput?.addEventListener("change", onFontFileSelected);
  els.addGoogleFontBtn?.addEventListener("click", addFontFromGoogleFonts);
  els.googleFontSearchInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      addFontFromGoogleFonts();
    }
  });
  els.resetAppFontBtn?.addEventListener("click", resetAppFont);

  els.defaultTextColorInput?.addEventListener("input", onDefaultTextColorLiveChange);
  els.wrapRuleColorInput?.addEventListener("input", onWrapRuleColorInputChange);
  els.wrapRuleHexInput?.addEventListener("input", onWrapRuleHexInputChange);
  els.wrapRuleSearchInput?.addEventListener("input", onWrapRuleSearchInput);

  els.novelAssistantImageBtn?.addEventListener("click", () => els.novelAssistantImageInput?.click());
  els.novelUserImageBtn?.addEventListener("click", () => els.novelUserImageInput?.click());
  els.novelChatImageBtn?.addEventListener("click", () => els.novelChatImageInput?.click());
  els.novelAssistantImageInput?.addEventListener("change", (event) => onNovelImageSelected(event, "assistant"));
  els.novelUserImageInput?.addEventListener("change", (event) => onNovelImageSelected(event, "user"));
  els.novelChatImageInput?.addEventListener("change", (event) => onNovelImageSelected(event, "chat"));

  els.exportDataBtn?.addEventListener("click", exportUserDataZip);
  els.importDataBtn?.addEventListener("click", () => els.importDataInput?.click());
  els.importDataInput?.addEventListener("change", importUserDataZip);

  els.drawerPresetSaveBtn?.addEventListener("click", saveDrawerPresetEdits);
  els.drawerPresetCancelBtn?.addEventListener("click", hideDrawerPresetEditor);
  els.drawerProviderSaveBtn?.addEventListener("click", saveDrawerProviderEdits);
  els.drawerProviderCancelBtn?.addEventListener("click", hideDrawerProviderEditor);
  els.drawerLogCopyBtn?.addEventListener("click", copySelectedLogToClipboard);
  els.drawerLogClearBtn?.addEventListener("click", clearPromptLogs);

  els.messageActionCloseBtn?.addEventListener("click", closeMessageActionPopup);
  els.messageActionPopup?.addEventListener("click", (event) => {
    if (event.target === els.messageActionPopup) closeMessageActionPopup();
  });
  els.deleteOneMessageBtn?.addEventListener("click", () => commitMessageDeleteFlow(false));
  els.deleteFromHereBtn?.addEventListener("click", () => commitMessageDeleteFlow(true));

  document.addEventListener("click", (event) => {
    const insideMessage = event.target?.closest?.(".message");
    if (!insideMessage) {
      if (expandedMessageMenuIndex !== -1) {
        expandedMessageMenuIndex = -1;
        renderMessages();
      }
      return;
    }
  });

  document.addEventListener("click", (event) => {
    const trigger = event.target?.closest?.("#textFontSectionBtn, #openFontManagerBtn");
    if (!trigger) return;
    event.preventDefault();
    openTextFontManager();
  });

  window.addEventListener("resize", updateComposerOffset);
}

function enterApp() {
  state.hasEnteredApp = true;
  persistState();
  applyWelcomeState();
}

function goHome() {
  toggleSettings(false);
  toggleChatDrawer(false);
  switchView("homeView");
  setStatus("Home");
}

function getHomeGreeting(date = new Date()) {
  const hour = date.getHours();
  if (hour < 12) return "Good morning,";
  if (hour < 18) return "Good afternoon,";
  return "Good evening,";
}

function updateHomeGreeting() {
  if (!els.homeGreeting) return;
  els.homeGreeting.textContent = getHomeGreeting();
}

function applyWelcomeState() {
  els.welcomeScreen.classList.toggle("hidden", state.hasEnteredApp);
  els.appShell.classList.toggle("hidden", !state.hasEnteredApp);
}

function switchView(viewId) {
  state.activeView = viewId;
  els.views.forEach((view) => view.classList.toggle("active", view.id === viewId));
  els.navCards.forEach((card) => {
    const isDirect = card.dataset.view === viewId;
    const isLorebookSubpage = viewId === "lorebookEditorView" && card.dataset.view === "libraryView";
    card.classList.toggle("active", isDirect || isLorebookSubpage);
  });
  if (viewId === "homeView") {
    updateHomeGreeting();
  }
  els.settingsToggle?.classList.add("hidden");
  if (viewId !== "chatsView") toggleChatDrawer(false);
  updateChatUiVisibility();
  syncLibraryLorebookFullscreenMode();
  persistState();
}

function toggleSettings(show) {
  els.settingsPanel.classList.toggle("hidden", !show);
}

function syncInputsFromState() {
  els.endpointInput.value = state.endpoint;
  els.apiKeyInput.value = state.apiKey;
  els.authHeaderInput.value = state.authHeader;
  els.authPrefixInput.value = state.authPrefix;
  els.modelsInput.value = state.models.join("\n");
  els.extraHeadersInput.value = state.extraHeaders;
  els.extraBodyInput.value = state.extraBody;
  if (els.systemPromptInput) {
    els.systemPromptInput.value = state.systemPrompt;
  }
  syncLorebookSettingsInputs();
  setInlineStatus(els.fetchModelsStatus, "");
  setInlineStatus(els.providerFetchModelsStatus, "");
}

function saveSettings() {
  const models = els.modelsInput.value
    .split(/\r?\n/)
    .map((m) => m.trim())
    .filter(Boolean);

  state.endpoint = els.endpointInput.value.trim();
  state.apiKey = els.apiKeyInput.value.trim();
  state.authHeader = els.authHeaderInput.value.trim() || defaults.authHeader;
  state.authPrefix = els.authPrefixInput.value;
  state.extraHeaders = els.extraHeadersInput.value.trim();
  state.extraBody = els.extraBodyInput.value.trim();
  state.models = models.length ? models : defaults.models;
  state.selectedModel = state.models.includes(state.selectedModel) ? state.selectedModel : state.models[0];
  state.systemPrompt = els.systemPromptInput?.value?.trim() || state.systemPrompt || defaults.systemPrompt;

  try {
    parseJsonObject(state.extraHeaders, "Extra headers");
    parseJsonObject(state.extraBody, "Extra body fields");
  } catch (error) {
    setStatus(error.message);
    return;
  }

  persistState();
  renderModelOptions();
  renderDrawerState();
  setStatus("Settings saved");
  toggleSettings(false);
}

async function clearStorage() {
  state = structuredClone(defaults);
  ensureCollections();
  syncInputsFromState();
  renderMessages();
  renderCharacters();
  renderProviders();
  renderPresets();
  renderLorebooks();
  renderRecentChats();
  renderImpersonationOptions();
  renderDrawerState();
  syncLorebookSettingsInputs();
  applyWelcomeState();
  switchView("homeView");
  maybeAutoFetchModels(true);

  if (hasFileSystemAccessApi()) {
    try {
      await writeUserFileToDisk(createUserFilePayload());
      setStatus(`Saved data cleared (wrote ${USER_FILE_NAME})`);
      showToast(`Cleared and wrote ${USER_FILE_NAME}`, "success");
      return;
    } catch {
      // fallback below
    }
  }

  persistState();
  setStatus("Saved data cleared");
}

function createChatRecord(title = "New chat") {
  const chat = {
    id: crypto.randomUUID(),
    title,
    characterId: "",
    lorebookIds: [],
    lorebookSettings: structuredClone(defaults.lorebookSettings),
    messages: [],
    preview: "",
    updatedAt: new Date().toISOString(),
  };
  state.chats.unshift(chat);
  state.currentChatId = chat.id;
  state.currentChatTitle = chat.title;
  return chat;
}

function ensureGreetingMessageForChat(chat) {
  if (!chat || !chat.characterId) return;
  const hasAnyMessage = Array.isArray(chat.messages) && chat.messages.length > 0;
  if (hasAnyMessage) return;

  const character = state.characters.find((item) => item.id === chat.characterId);
  if (!character) return;

  const greetingVariants = getCharacterGreetingVariants(character)
    .map((value) => applyMacrosToText(value, character))
    .map((value) => value.trim())
    .filter(Boolean);

  if (chat.messages?.length && chat.messages[0]?.isGreeting && Array.isArray(chat.messages[0]?.swipes)) {
    chat.messages[0].swipes = greetingVariants;
    chat.messages[0].swipe_id = clampSwipeIndex(chat.messages[0].swipe_id, greetingVariants.length);
    chat.messages[0].content = greetingVariants[chat.messages[0].swipe_id] || greetingVariants[0] || "";
    chat.messages[0].swipeable = greetingVariants.length > 1;
    return;
  }


  if (!greetingVariants.length) return;

  chat.messages = [{
    role: "assistant",
    content: greetingVariants[0],
    swipes: greetingVariants,
    swipe_id: 0,
    swipeable: greetingVariants.length > 1,
    isGreeting: true,
  }];
}

function getCurrentChat() {
  return state.chats.find((chat) => chat.id === state.currentChatId) || null;
}

function ensureCurrentChat() {
  return getCurrentChat() || createChatRecord(state.currentChatTitle || "New chat");
}

function refreshRecentChatsFromState() {
  state.recentChats = state.chats
    .slice()
    .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))
    .slice(0, 8)
    .map(({ id, title, preview, updatedAt, characterId }) => ({ id, title, preview, updatedAt, characterId }));
}

function syncCurrentChatFromState() {
  const chat = getCurrentChat();
  if (!chat) {
    refreshRecentChatsFromState();
    return;
  }
  chat.title = state.currentChatTitle || chat.title || "New chat";
  chat.messages = state.messages.map((message) => ({ ...message }));
  chat.lorebookIds = Array.isArray(chat.lorebookIds) ? [...new Set(chat.lorebookIds.map((id) => String(id || "").trim()).filter(Boolean))] : [];
  chat.lorebookSettings = {
    ...structuredClone(defaults.lorebookSettings),
    ...(chat.lorebookSettings && typeof chat.lorebookSettings === "object" ? chat.lorebookSettings : {}),
  };
  chat.lorebookSettings.scanDepth = Math.min(64, Math.max(1, Number.parseInt(chat.lorebookSettings.scanDepth, 10) || defaults.lorebookSettings.scanDepth));
  chat.lorebookSettings.tokenBudget = Math.min(8000, Math.max(64, Number.parseInt(chat.lorebookSettings.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget));
  chat.lorebookSettings.recursiveScanning = chat.lorebookSettings.recursiveScanning !== false;

  const previewSource = [...chat.messages].reverse().find((message) => message.role === "assistant" || message.role === "user");
  chat.preview = (previewSource?.content || "").slice(0, 80);
  chat.updatedAt = new Date().toISOString();
  refreshRecentChatsFromState();
}

function openLastInteractedChat() {
  const current = state.currentChatId
    ? state.chats.find((chat) => chat.id === state.currentChatId)
    : null;
  if (current) {
    openChat(current.id);
    return;
  }

  refreshRecentChatsFromState();
  const latest = state.recentChats[0];
  if (latest?.id) {
    openChat(latest.id);
    return;
  }

  switchView("chatsView");
  state.messages = [];
  state.currentChatId = "";
  state.currentChatTitle = "New chat";
  renderMessages();
  renderRecentChats();
  renderDrawerState();
  updateChatUiVisibility();
  setStatus("No chats yet. Start one from a character.");
}

async function onSend(event) {
  event?.preventDefault?.();
  const currentCharacter = getCurrentChatCharacter();
  const content = applyMacrosToText(els.messageInput.value.trim(), currentCharacter);
  if (!content) return;

  if (!state.endpoint || !state.apiKey) {
    setStatus("Add endpoint and API key in settings");
    toggleSettings(true);
    return;
  }

  switchView("chatsView");
  ensureCurrentChat();

  const userMessage = { role: "user", content };
  state.messages.push(userMessage);
  state.lastGenerationType = "user";

  els.messageInput.value = "";
  autoResizeTextarea(els.messageInput);

  const assistantMessage = { role: "assistant", content: "", pending: true };
  state.messages.push(assistantMessage);

  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();

  const requestMessages = state.messages.slice(0, -1);
  await generateAssistantReply({ requestMessages, assistantMessage });
}

function extractThinkingSummaryFromText(text = "", provider = null) {
  const source = String(text || "");
  const startTag = provider?.thinkingStartTag || "<thinking>";
  const endTag = provider?.thinkingEndTag || "</thinking>";
  if (!startTag || !endTag) return "";
  const escapedStart = escapeRegex(startTag);
  const escapedEnd = escapeRegex(endTag);
  const regex = new RegExp(`${escapedStart}([\\s\\S]*?)${escapedEnd}`, "gi");
  const chunks = [];
  let match;
  while ((match = regex.exec(source)) !== null) {
    chunks.push(String(match[1] || "").trim());
  }
  return chunks.filter(Boolean).join("\n").trim();
}

function stripThinkingBlocks(text = "", provider = null) {
  const source = String(text || "");
  const startTag = provider?.thinkingStartTag || "<thinking>";
  const endTag = provider?.thinkingEndTag || "</thinking>";
  if (!startTag || !endTag) return source;
  const escapedStart = escapeRegex(startTag);
  const escapedEnd = escapeRegex(endTag);
  return source.replace(new RegExp(`${escapedStart}[\\s\\S]*?${escapedEnd}`, "gi"), "").trim();
}

function getActiveProvider() {
  return state.providers.find((item) => item.id === state.activeProviderId) || null;
}

function appendPromptLogEntry({ payload = {}, mode = "standard" } = {}) {
  const endpoint = state.endpoint ? normalizeEndpoint(state.endpoint) : "";
  const entry = {
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    mode,
    model: String(payload?.model || state.selectedModel || ""),
    endpoint,
    chatId: state.currentChatId || "",
    chatTitle: state.currentChatTitle || "",
    payload: payload && typeof payload === "object" ? JSON.parse(JSON.stringify(payload)) : {},
  };

  state.promptLogs.unshift(entry);
  state.promptLogs = state.promptLogs.slice(0, 5);

  if (!drawerSelectedLogId || !state.promptLogs.some((log) => log.id === drawerSelectedLogId)) {
    drawerSelectedLogId = entry.id;
  }

  if (els.chatDrawer?.dataset?.panel === "log") {
    renderDrawerLogs();
  }
}

function getSelectedPromptLog() {
  if (!Array.isArray(state.promptLogs) || !state.promptLogs.length) return null;
  return state.promptLogs.find((entry) => entry.id === drawerSelectedLogId) || state.promptLogs[0] || null;
}

function formatPromptLogTerminalEntry(entry) {
  if (!entry) return "";
  const lines = [
    `$ maitavern prompt-log --id ${entry.id}`,
    `$ created_at: ${entry.createdAt}`,
    `$ mode: ${entry.mode || "standard"}`,
    `$ endpoint: ${entry.endpoint || "(not set)"}`,
    `$ model: ${entry.model || "(unknown)"}`,
    `$ chat: ${entry.chatTitle || "New chat"}${entry.chatId ? ` (${entry.chatId})` : ""}`,
    "$ payload:",
    JSON.stringify(entry.payload || {}, null, 2),
  ];
  return lines.join("\n");
}

function renderDrawerLogs() {
  if (!els.drawerLogList || !els.drawerLogTerminal) return;

  const logs = Array.isArray(state.promptLogs) ? state.promptLogs : [];
  if (!drawerSelectedLogId || !logs.some((entry) => entry.id === drawerSelectedLogId)) {
    drawerSelectedLogId = logs[0]?.id || "";
  }

  els.drawerLogList.innerHTML = "";

  if (!logs.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No prompt logs yet. Send a message to capture request payloads.</div>`;
    els.drawerLogList.appendChild(empty);
    if (els.drawerLogMeta) els.drawerLogMeta.textContent = "Logs: 0/5";
    els.drawerLogTerminal.textContent = "$ waiting for prompt logs...";
    return;
  }

  logs.forEach((entry, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `drawer-choice${entry.id === drawerSelectedLogId ? " active" : ""}`;
    const timeLabel = new Date(entry.createdAt || Date.now()).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    button.innerHTML = `
      <strong>#${logs.length - index} • ${escapeHtml(entry.model || "model")}</strong>
      <small>${escapeHtml(timeLabel)} • ${escapeHtml(entry.mode || "standard")}</small>
    `;
    button.addEventListener("click", () => {
      drawerSelectedLogId = entry.id;
      renderDrawerLogs();
    });
    els.drawerLogList.appendChild(button);
  });

  const selected = getSelectedPromptLog();
  if (els.drawerLogMeta) {
    els.drawerLogMeta.textContent = selected
      ? `Logs: ${logs.length}/5 • Showing ${selected.model || "unknown model"}`
      : `Logs: ${logs.length}/5`;
  }
  els.drawerLogTerminal.textContent = formatPromptLogTerminalEntry(selected);
}

async function copySelectedLogToClipboard() {
  const selected = getSelectedPromptLog();
  if (!selected) {
    showToast("No log selected", "error");
    return;
  }

  const text = formatPromptLogTerminalEntry(selected);

  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      const helper = document.createElement("textarea");
      helper.value = text;
      helper.style.position = "fixed";
      helper.style.opacity = "0";
      document.body.appendChild(helper);
      helper.focus();
      helper.select();
      document.execCommand("copy");
      helper.remove();
    }
    showToast("Log copied", "success");
  } catch {
    showToast("Could not copy log", "error");
  }
}

function clearPromptLogs() {
  state.promptLogs = [];
  drawerSelectedLogId = "";
  persistState();
  renderDrawerLogs();
  showToast("Prompt logs cleared", "success");
}

async function generateAssistantReply({ requestMessages, assistantMessage }) {
  setBusy(true);
  setStatus(`Calling ${state.selectedModel}...`);
  currentController = new AbortController();

  try {
    const extraHeaders = parseJsonObject(state.extraHeaders, "Extra headers");
    const extraBody = parseJsonObject(state.extraBody, "Extra body fields");
    const headers = {
      "Content-Type": "application/json",
      ...extraHeaders,
    };

    if (state.apiKey && state.authHeader) {
      headers[state.authHeader] = `${state.authPrefix || ""}${state.apiKey}`;
    }

    const preset = getActivePreset();
    const messagesForRequest = buildApiMessages(requestMessages);
    const messagesForLog = JSON.parse(JSON.stringify(messagesForRequest));

    const basePayload = {
      model: state.selectedModel,
      messages: messagesForRequest,
      stream: Boolean(preset?.stream),
      ...extraBody,
    };

    const basePayloadForLog = {
      model: state.selectedModel,
      messages: messagesForLog,
      stream: Boolean(preset?.stream),
      ...extraBody,
    };

    const standardPayload = {
      ...basePayload,
      temperature: preset.temperature ?? 0.7,
      top_p: preset.topP ?? 1,
      max_tokens: preset.maxTokens ?? 300,
      ...(shouldApplyInstructPreset(preset) && preset.instruct?.stopSequence
        ? { stop: [preset.instruct.stopSequence] }
        : {}),
    };

    const standardPayloadForLog = {
      ...basePayloadForLog,
      temperature: preset.temperature ?? 0.7,
      top_p: preset.topP ?? 1,
      max_tokens: preset.maxTokens ?? 300,
      ...(shouldApplyInstructPreset(preset) && preset.instruct?.stopSequence
        ? { stop: [preset.instruct.stopSequence] }
        : {}),
    };

    appendPromptLogEntry({ payload: standardPayloadForLog, mode: "standard" });

    let data = await requestChatCompletion({ headers, payload: standardPayload, signal: currentController.signal });
    let text = extractAssistantText(data);

    if (!text && shouldRetryWithCompatPayload(data)) {
      const compatPayload = {
        ...basePayload,
        max_completion_tokens: preset.maxTokens ?? 300,
        ...(shouldApplyInstructPreset(preset) && preset.instruct?.stopSequence
          ? { stop: [preset.instruct.stopSequence] }
          : {}),
      };
      const compatPayloadForLog = {
        ...basePayloadForLog,
        max_completion_tokens: preset.maxTokens ?? 300,
        ...(shouldApplyInstructPreset(preset) && preset.instruct?.stopSequence
          ? { stop: [preset.instruct.stopSequence] }
          : {}),
      };
      appendPromptLogEntry({ payload: compatPayloadForLog, mode: "compat" });
      data = await requestChatCompletion({ headers, payload: compatPayload, signal: currentController.signal });
      text = extractAssistantText(data);
    }

    if (!text) {
      console.warn("MaiTavern: empty or unrecognized response payload", data);
    }

    const activeProvider = getActiveProvider();
    const rawText = text || `Provider returned an empty assistant message. ${JSON.stringify(data).slice(0, 280)}`;
    assistantMessage.pending = false;
    assistantMessage.thinkingSummary = extractThinkingSummaryFromText(rawText, activeProvider);
    assistantMessage.content = stripThinkingBlocks(rawText, activeProvider);
    state.lastGenerationType = "assistant";

    syncCurrentChatFromState();
    persistState();
    renderMessages();
    renderRecentChats();
    setStatus(text ? "Done" : "Provider returned empty content");
  } catch (error) {
    assistantMessage.pending = false;
    assistantMessage.content = `Error: ${error.message}`;
    syncCurrentChatFromState();
    persistState();
    renderMessages();
    renderRecentChats();
    setStatus(error.name === "AbortError" ? "Request stopped" : "Request failed");
  } finally {
    setBusy(false);
    currentController = null;
  }
}

function stopRequest() {
  currentController?.abort();
}

function buildApiMessages(messages, options = {}) {
  const preset = getActivePreset();
  const currentCharacter = getCurrentChatCharacter();
  const limitConversationMessages = Number(options?.limitConversationMessages);

  const activeProvider = getActiveProvider();
  let filteredMessages = messages
    .filter((msg) => msg.role === "user" || msg.role === "assistant")
    .map((msg) => {
      if (activeProvider?.includeThinkingInPrompt && msg.thinkingSummary) {
        return {
          ...msg,
          content: `${msg.content}\n\n[Thinking summary]\n${msg.thinkingSummary}`,
        };
      }
      return msg;
    });

  if (Number.isFinite(limitConversationMessages) && limitConversationMessages > 0) {
    filteredMessages = filteredMessages.slice(-limitConversationMessages);
  }

  const loreContext = resolveLorebookPromptContext({ messages: filteredMessages, character: currentCharacter });
  const systemPrompt = buildPresetSystemPrompt({ preset, character: currentCharacter, loreContext });
  const vars = buildPromptVariables(currentCharacter, systemPrompt, loreContext);

  const injected = buildPresetInjectedMessages({ preset, vars });

  if (shouldApplyInstructPreset(preset)) {
    const injectedSystemBlocks = [...injected.before, ...injected.after]
      .filter((msg) => msg.role === "system")
      .map((msg) => String(msg.content || "").trim())
      .filter(Boolean);

    const instructSystemPrompt = [systemPrompt, ...injectedSystemBlocks]
      .filter(Boolean)
      .join("\n\n")
      .trim();

    const instructMessages = [
      ...injected.before.filter((msg) => msg.role !== "system"),
      ...filteredMessages.map((msg) => ({ role: msg.role, content: msg.content })),
      ...injected.after.filter((msg) => msg.role !== "system"),
    ];

    return buildInstructTranscriptMessages({ messages: instructMessages, preset, systemPrompt: instructSystemPrompt });
  }

  const apiMessages = [];

  if (systemPrompt) {
    apiMessages.push({ role: "system", content: systemPrompt });
  }

  injected.before.forEach((msg) => apiMessages.push(msg));

  for (const msg of filteredMessages) {
    apiMessages.push({ role: msg.role, content: msg.content });
  }

  injected.after.forEach((msg) => apiMessages.push(msg));

  return apiMessages;
}

function shouldApplyInstructPreset(preset) {
  if (typeof preset?.instruct?.enabled === "boolean") {
    return preset.instruct.enabled;
  }
  return preset?.source === "sillytavern-instruct";
}

function buildInstructTranscriptMessages({ messages, preset, systemPrompt }) {
  const transcript = buildInstructTranscript({ messages, preset, systemPrompt });
  return [{ role: "user", content: transcript }];
}

function buildInstructTranscript({ messages, preset, systemPrompt }) {
  const instruct = preset?.instruct || {};
  const transcriptParts = [];
  const dialogue = messages.filter((msg) => msg.role === "user" || msg.role === "assistant");

  if (systemPrompt) {
    transcriptParts.push(formatInstructMessageContent({
      role: "system",
      content: systemPrompt,
      index: 0,
      total: dialogue.length + 1,
      instruct,
    }));
  }

  dialogue.forEach((message, idx) => {
    transcriptParts.push(formatInstructMessageContent({
      role: message.role,
      content: message.content,
      index: idx + 1,
      total: dialogue.length + 1,
      instruct,
    }));
  });

  const assistantCue = buildAssistantContinuationCue(instruct);
  if (assistantCue) {
    transcriptParts.push(assistantCue);
  }

  return transcriptParts.filter(Boolean).join(instruct.wrap ? "\n" : "\n\n").trim();
}

function buildAssistantContinuationCue(instruct) {
  const sequenceRaw = instruct.lastOutputSequence || instruct.outputSequence || "";
  const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
  if (!sequence) return "";
  return `${sequence}${sequence ? "\n" : ""}`;
}

function formatInstructMessageContent({ role, content, index, total, instruct }) {
  const isFirstNonSystem = index === 1;
  const isLast = index === total - 1;
  const stopSequence = applyMacrosToText(instruct.stopSequence || "", getCurrentChatCharacter());
  const resolvedContent = applyMacrosToText(String(content || ""), getCurrentChatCharacter());

  if (role === "system") {
    const sequenceRaw = instruct.systemSameAsUser
      ? (instruct.firstInputSequence || instruct.inputSequence || "")
      : (instruct.systemSequence || instruct.storyStringPrefix || "");
    const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
    const suffix = applyMacrosToText(instruct.systemSuffix || instruct.storyStringSuffix || stopSequence || "", getCurrentChatCharacter());
    return `${sequence}${sequence ? "\n" : ""}${resolvedContent}${suffix}`.trimEnd();
  }

  if (role === "user") {
    const sequenceRaw = isFirstNonSystem
      ? (instruct.firstInputSequence || instruct.inputSequence || "")
      : isLast
        ? (instruct.lastInputSequence || instruct.inputSequence || "")
        : (instruct.inputSequence || "");
    const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
    const suffix = applyMacrosToText(instruct.inputSuffix || stopSequence || "", getCurrentChatCharacter());
    return `${sequence}${sequence ? "\n" : ""}${resolvedContent}${suffix}`.trimEnd();
  }

  if (role === "assistant") {
    const sequenceRaw = isFirstNonSystem
      ? (instruct.firstOutputSequence || instruct.outputSequence || "")
      : isLast
        ? (instruct.lastOutputSequence || instruct.outputSequence || "")
        : (instruct.outputSequence || "");
    const sequence = applyMacrosToText(sequenceRaw, getCurrentChatCharacter());
    const suffix = applyMacrosToText(instruct.outputSuffix || stopSequence || "", getCurrentChatCharacter());
    return `${sequence}${sequence ? "\n" : ""}${resolvedContent}${suffix}`.trimEnd();
  }

  return resolvedContent;
}

function buildPresetSystemPrompt({ preset, character, loreContext = null }) {
  const includeSystemPrompt = preset?.useSystemPrompt !== false;
  const fallbackPrompt = includeSystemPrompt ? (preset?.systemPrompt?.trim() || state.systemPrompt?.trim() || "") : "";
  const vars = buildPromptVariables(character, fallbackPrompt, loreContext);
  const contextBlock = buildContextPresetPrompt({ preset, vars });
  const template = preset?.contextTemplate || defaults.presets[0].contextTemplate;
  const templateHandlesLore = /{{\s*(wiBefore|wiAfter|anchorBefore|anchorAfter)\s*}}/i.test(template);

  const parts = [fallbackPrompt];
  if (!templateHandlesLore && vars.wiBefore) parts.push(vars.wiBefore);
  if (contextBlock) parts.push(contextBlock);
  if (!templateHandlesLore && vars.wiAfter) parts.push(vars.wiAfter);

  return parts.filter(Boolean).join("\n\n").trim();
}

function buildPresetInjectedMessages({ preset, vars }) {
  const before = [];
  const after = [];

  getOrderedPromptEntries(preset)
    .filter((prompt) => prompt._orderEnabled !== false && prompt.enabled !== false)
    .forEach((prompt) => {
      const rendered = renderOpenAiPromptEntry(prompt, vars);
      if (!rendered) return;
      const role = ["system", "user", "assistant"].includes(prompt.role) ? prompt.role : "system";
      const target = Number(prompt.injection_position ?? 0) === 1 ? after : before;
      target.push({ role, content: rendered });
    });

  return { before, after };
}

function buildPromptVariables(character, systemPrompt = "", loreContext = null) {
  const name = character?.name?.trim() || "Character";
  const userName = resolveCurrentUserName();
  const preset = getActivePreset();
  const scenario = character?.scenario?.trim() || character?.data?.scenario?.trim() || "";
  const personality = character?.personality?.trim() || character?.data?.personality?.trim() || character?.data?.description?.trim() || "";
  const greeting = character?.greeting?.trim() || character?.data?.first_mes?.trim() || "";
  const description = character?.data?.description?.trim() || personality || "";
  const creatorNotes = character?.data?.creator_notes || character?.data?.creatorNotes || "";
  const charVersion = character?.data?.character_version || character?.data?.char_version || character?.data?.version || "";
  const mesExamplesRaw = character?.data?.mes_example || character?.data?.mesExamples || "";
  const loreBeforeRaw = String(loreContext?.beforeText || "").trim();
  const loreAfterRaw = String(loreContext?.afterText || "").trim();

  const vars = {
    char: name,
    user: userName,
    system: systemPrompt,
    systemPrompt,
    defaultSystemPrompt: systemPrompt,
    instructSystem: systemPrompt,
    instructSystemPrompt: systemPrompt,
    description,
    personality,
    scenario,
    persona: "",
    anchorBefore: loreBeforeRaw,
    anchorAfter: loreAfterRaw,
    wiBefore: loreBeforeRaw,
    wiAfter: loreAfterRaw,
    greeting,
    charDescription: description,
    charPersonality: personality,
    charScenario: scenario,
    charFirstMessage: greeting,
    charCreatorNotes: creatorNotes,
    creatorNotes,
    charVersion,
    version: charVersion,
    char_version: charVersion,
    mesExamplesRaw,
    mesExamples: mesExamplesRaw,
    chatStart: preset?.chatStart || "",
    exampleSeparator: preset?.exampleSeparator || "",
    chatSeparator: preset?.exampleSeparator || "",
  };

  const resolvedGreeting = applyMacrosToText(greeting, character, userName);
  const resolvedWiBefore = applyMacrosToText(loreBeforeRaw, character, userName, { vars });
  const resolvedWiAfter = applyMacrosToText(loreAfterRaw, character, userName, { vars });

  return {
    ...vars,
    greeting: resolvedGreeting,
    wiBefore: resolvedWiBefore,
    wiAfter: resolvedWiAfter,
    anchorBefore: resolvedWiBefore,
    anchorAfter: resolvedWiAfter,
  };
}

function getOrderedPromptEntries(preset) {
  const promptOrder = Array.isArray(preset?.promptOrder) ? preset.promptOrder : [];
  const prompts = normalizePresetPromptList(Array.isArray(preset?.prompts) ? preset.prompts : []);

  const orderBlock = promptOrder.find((item) => Array.isArray(item?.order));
  if (!orderBlock) {
    return prompts.map((prompt) => ({ ...prompt, _orderEnabled: prompt.enabled !== false }));
  }

  const map = new Map(prompts.map((prompt) => [prompt.identifier, prompt]));
  const ordered = [];

  orderBlock.order.forEach((entry) => {
    const prompt = map.get(entry?.identifier);
    if (!prompt) return;
    ordered.push({ ...prompt, _orderEnabled: entry?.enabled !== false });
  });

  prompts.forEach((prompt) => {
    if (ordered.some((item) => item.identifier === prompt.identifier)) return;
    ordered.push({ ...prompt, _orderEnabled: prompt.enabled !== false });
  });

  return ordered;
}

function buildOpenAiPresetPrompt({ preset, vars, positionFilter = null }) {
  const ordered = getOrderedPromptEntries(preset)
    .filter((prompt) => prompt._orderEnabled !== false && prompt.enabled !== false)
    .filter((prompt) => {
      if (positionFilter === null) return true;
      return Number(prompt.injection_position ?? 0) === Number(positionFilter);
    });

  const rendered = ordered
    .map((prompt) => renderOpenAiPromptEntry(prompt, vars))
    .filter(Boolean);

  return rendered.join("\n\n").trim();
}

function renderOpenAiPromptEntry(prompt, vars) {
  const identifier = prompt.identifier || "";
  if (prompt.marker) {
    if (identifier === "charDescription") return renderTemplate(vars.description, vars);
    if (identifier === "charPersonality") return renderTemplate(vars.personality, vars);
    if (identifier === "scenario") return renderTemplate(vars.scenario, vars);
    if (identifier === "personaDescription") return renderTemplate(vars.persona, vars);
    if (identifier === "main") return renderTemplate(vars.system, vars);
    return "";
  }
  return renderTemplate(prompt.content || "", vars);
}

function buildContextPresetPrompt({ preset, vars }) {
  const template = preset?.contextTemplate || defaults.presets[0].contextTemplate;
  const formattedScenario = preset?.scenarioFormat ? renderTemplate(preset.scenarioFormat, vars) : vars.scenario;
  const formattedPersonality = preset?.personalityFormat ? renderTemplate(preset.personalityFormat, vars) : vars.personality;
  return renderTemplate(template, {
    ...vars,
    scenario: formattedScenario,
    personality: formattedPersonality,
  }).trim();
}

function renderTemplate(template, vars) {
  let output = String(template || "");
  output = output.replace(/{{#if\s+([a-zA-Z0-9_]+)}}([\s\S]*?){{\/if}}/g, (_, key, content) => vars?.[key] ? content : "");
  Object.entries(vars || {}).forEach(([key, value]) => {
    output = output.replace(new RegExp(`{{\\s*${escapeRegex(key)}\\s*}}`, "g"), String(value || ""));
  });
  output = applyMacrosToText(output, null, vars?.user || "", { vars });
  output = output.replace(/\n{3,}/g, "\n\n");
  return output.trim();
}

function estimateTokenCount(text = "") {
  const normalized = String(text || "").trim();
  if (!normalized) return 0;
  return Math.ceil(normalized.length / 4);
}

function normalizeStringArray(value) {
  const output = [];

  const append = (item) => {
    if (item === undefined || item === null) return;

    if (Array.isArray(item)) {
      item.forEach(append);
      return;
    }

    if (typeof item === "string") {
      item
        .split(/[\n,]/)
        .map((part) => part.trim())
        .filter(Boolean)
        .forEach((part) => output.push(part));
      return;
    }

    if (typeof item === "number") {
      const valueText = String(item).trim();
      if (valueText) output.push(valueText);
      return;
    }

    if (typeof item === "object") {
      append(
        item.keyword
        ?? item.key
        ?? item.name
        ?? item.label
        ?? item.value
        ?? item.text
        ?? item.word
        ?? ""
      );
    }
  };

  append(value);

  return [...new Set(output.map((item) => String(item || "").trim()).filter(Boolean))];
}

function normalizeSelectiveLogic(value = "AND") {
  const raw = String(value ?? "AND").trim().toUpperCase();
  if (raw === "0") return "AND";
  if (raw === "1") return "OR";
  if (raw === "2") return "NOT";
  if (["AND", "OR", "NOT"].includes(raw)) return raw;
  return "AND";
}

function normalizeInsertionPosition(value) {
  const raw = value ?? "after_system";

  if (typeof raw === "number" && Number.isFinite(raw)) {
    const map = {
      0: "before_system",
      1: "after_system",
      2: "before_author_note",
      3: "after_author_note",
      4: "top_chat",
      5: "bottom_chat",
      6: "depth",
    };
    return map[raw] || "after_system";
  }

  const normalized = String(raw || "").trim().toLowerCase();
  const map = {
    before_system: "before_system",
    before_system_prompt: "before_system",
    before_prompt: "before_system",
    after_system: "after_system",
    after_system_prompt: "after_system",
    before_author_note: "before_author_note",
    after_author_note: "after_author_note",
    top_chat: "top_chat",
    top_of_chat: "top_chat",
    bottom_chat: "bottom_chat",
    bottom_of_chat: "bottom_chat",
    depth: "depth",
    "0": "before_system",
    "1": "after_system",
    "2": "before_author_note",
    "3": "after_author_note",
    "4": "top_chat",
    "5": "bottom_chat",
    "6": "depth",
  };

  return map[normalized] || "after_system";
}

function looksLikeLorebookEntryObject(item = {}) {
  if (!item || typeof item !== "object") return false;
  return (
    Object.prototype.hasOwnProperty.call(item, "content")
    || Object.prototype.hasOwnProperty.call(item, "entry")
    || Object.prototype.hasOwnProperty.call(item, "text")
    || Object.prototype.hasOwnProperty.call(item, "value")
    || Object.prototype.hasOwnProperty.call(item, "description")
    || Object.prototype.hasOwnProperty.call(item, "key")
    || Object.prototype.hasOwnProperty.call(item, "keys")
    || Object.prototype.hasOwnProperty.call(item, "keywords")
    || Object.prototype.hasOwnProperty.call(item, "primary_keys")
    || Object.prototype.hasOwnProperty.call(item, "secondary_keys")
    || Object.prototype.hasOwnProperty.call(item, "secondaryKeywords")
    || Object.prototype.hasOwnProperty.call(item, "keysecondary")
    || Object.prototype.hasOwnProperty.call(item, "match")
    || Object.prototype.hasOwnProperty.call(item, "triggers")
  );
}

function normalizeLorebookEntry(entry = {}, index = 0) {
  const keys = normalizeStringArray(
    entry.keys
    || entry.key
    || entry.keywords
    || entry.primary_keys
    || entry.triggers
    || entry.match
    || []
  );
  const secondaryKeys = normalizeStringArray(
    entry.secondary_keys
    || entry.secondaryKeys
    || entry.secondary
    || entry.secondary_keywords
    || entry.secondaryKeywords
    || entry.keysecondary
    || entry.key_secondary
    || []
  );

  const content = String(
    entry.content
    ?? entry.text
    ?? entry.value
    ?? entry.entry
    ?? entry.prompt
    ?? entry.desc
    ?? entry.description
    ?? ""
  ).trim();

  const insertionOrder = Number.isFinite(Number(entry.insertion_order ?? entry.insertionOrder ?? entry.order ?? entry.uid))
    ? Number(entry.insertion_order ?? entry.insertionOrder ?? entry.order ?? entry.uid)
    : index;

  const priority = Number.isFinite(Number(entry.priority ?? entry.weight ?? entry.extensions?.weight))
    ? Number(entry.priority ?? entry.weight ?? entry.extensions?.weight)
    : Number.isFinite(Number(entry.insertionPriority))
      ? Number(entry.insertionPriority)
      : 100;

  const probabilityRaw = Number(
    entry.probability
    ?? entry.chance
    ?? entry.activationProbability
    ?? entry.extensions?.probability
    ?? 100
  );
  const probability = Math.min(100, Math.max(0, Number.isFinite(probabilityRaw) ? probabilityRaw : 100));

  const selectiveLogic = normalizeSelectiveLogic(
    entry.selectiveLogic
    ?? entry.selective_logic
    ?? entry.logic
    ?? entry.extensions?.selectiveLogic
    ?? "AND"
  );

  let activationLogic = String(entry.activationLogic || entry.activation_logic || "").trim().toUpperCase();
  if (!activationLogic) {
    if (Boolean(entry.selective ?? secondaryKeys.length > 0)) {
      if (selectiveLogic === "NOT") activationLogic = "NOT";
      else if (selectiveLogic === "AND") activationLogic = "AND_ANY_SECONDARY";
      else activationLogic = "ANY";
    } else {
      activationLogic = "ANY";
    }
  }
  if (!["ANY", "ALL", "AND_ANY_SECONDARY", "NOT"].includes(activationLogic)) {
    activationLogic = "ANY";
  }

  const strategyRaw = String(entry.strategy || entry.activationStrategy || "").trim().toLowerCase();
  const strategy = ["normal", "constant", "vectorized"].includes(strategyRaw)
    ? strategyRaw
    : (entry.constant ? "constant" : "normal");

  const nowIso = new Date().toISOString();

  const positionSource = entry.insertionPosition
    ?? entry.insertion_position
    ?? entry.position
    ?? entry.extensions?.position
    ?? "after_system";

  const caseSource = entry.caseSensitive
    ?? entry.case_sensitive
    ?? entry.caseSensitiveMatching
    ?? entry.case_sensitive_matching
    ?? false;

  return {
    id: entry.id || entry.uid || crypto.randomUUID(),
    title: String(entry.title || entry.name || entry.comment || `Entry ${index + 1}`),
    comment: String(entry.comment || entry.note || entry.memo || ""),
    keys,
    secondaryKeys,
    activationLogic,
    matchWholeWords: Boolean(entry.matchWholeWords ?? entry.match_whole_words ?? false),
    useRegex: Boolean(entry.useRegex ?? entry.use_regex ?? false),
    content,
    insertionOrder,
    priority,
    insertionPosition: normalizeInsertionPosition(positionSource),
    insertionDepth: Math.max(
      0,
      Number.parseInt(
        entry.insertionDepth
        ?? entry.insertion_depth
        ?? entry.depth
        ?? entry.extensions?.depth
        ?? 4,
        10
      ) || 4
    ),
    role: ["system", "user", "assistant"].includes(String(entry.role || "system").toLowerCase())
      ? String(entry.role || "system").toLowerCase()
      : "system",
    caseSensitive: Boolean(caseSource),
    selective: Boolean(entry.selective ?? secondaryKeys.length > 0),
    selectiveLogic,
    scanDepth: Math.max(0, Number.parseInt(entry.scanDepth ?? entry.scan_depth ?? 0, 10) || 0),
    tokenBudget: Math.max(0, Number.parseInt(entry.tokenBudget ?? entry.token_budget ?? 0, 10) || 0),
    constant: strategy === "constant" || Boolean(entry.constant ?? false),
    strategy,
    probability,
    enabled: entry.enabled !== false && entry.disable !== true,
    preventRecursion: Boolean(entry.preventRecursion ?? entry.prevent_recursion ?? false),
    excludeFromRecursion: Boolean(
      entry.excludeFromRecursion
      ?? entry.exclude_from_recursion
      ?? entry.excludeRecursion
      ?? entry.exclude_recursion
      ?? entry.extensions?.excludeRecursion
      ?? false
    ),
    group: String(entry.group || ""),
    groupWeight: Math.max(1, Number.parseInt(entry.groupWeight ?? entry.group_weight ?? 100, 10) || 100),
    category: String(entry.category || ""),
    tags: normalizeStringArray(entry.tags || entry.tag || []),
    automationId: String(entry.automationId || entry.automation_id || ""),
    linkedEntries: normalizeStringArray(entry.linkedEntries || entry.linked_entries || []),
    sticky: Math.max(0, Number.parseInt(entry.sticky ?? entry.stickyDuration ?? entry.sticky_duration ?? 0, 10) || 0),
    cooldown: Math.max(0, Number.parseInt(entry.cooldown ?? 0, 10) || 0),
    delay: Math.max(0, Number.parseInt(entry.delay ?? 0, 10) || 0),
    turnStart: Math.max(0, Number.parseInt(entry.turnStart ?? entry.turn_start ?? 0, 10) || 0),
    turnEnd: Number.isFinite(Number.parseInt(entry.turnEnd ?? entry.turn_end, 10))
      ? Math.max(0, Number.parseInt(entry.turnEnd ?? entry.turn_end, 10))
      : null,
    createdAt: String(entry.createdAt || entry.created_at || nowIso),
    updatedAt: String(entry.updatedAt || entry.updated_at || nowIso),
  };
}

function parseLorebookEntries(raw = {}) {
  if (!raw || typeof raw !== "object") return [];

  const candidateSources = [
    raw.entries,
    raw.data?.entries,
    raw.world_info,
    raw.world_info?.entries,
    raw.worldInfo,
    raw.worldInfo?.entries,
    raw.lorebook?.entries,
    raw.book?.entries,
    raw.items,
    raw.memory?.entries,
    raw.data?.memory?.entries,
  ];

  const seen = new Set();
  const entriesSource = [];

  const addSource = (source) => {
    if (!source) return;

    if (Array.isArray(source)) {
      source.forEach((item) => {
        if (!item || typeof item !== "object") return;
        if (seen.has(item)) return;
        seen.add(item);
        entriesSource.push(item);
      });
      return;
    }

    if (typeof source === "object") {
      Object.values(source).forEach((item) => {
        if (!item || typeof item !== "object") return;
        if (seen.has(item)) return;
        seen.add(item);
        entriesSource.push(item);
      });
    }
  };

  candidateSources.forEach(addSource);

  if (!entriesSource.length) {
    const queue = [raw];
    const scanned = new Set();

    while (queue.length) {
      const current = queue.shift();
      if (!current || typeof current !== "object") continue;
      if (scanned.has(current)) continue;
      scanned.add(current);

      if (current !== raw && looksLikeLorebookEntryObject(current)) {
        addSource([current]);
      }

      Object.values(current).forEach((value) => {
        if (value && typeof value === "object") queue.push(value);
      });
    }
  }

  return entriesSource
    .map((entry, index) => normalizeLorebookEntry(entry, index))
    .filter((entry) => {
      const hasKeys = Array.isArray(entry.keys) && entry.keys.length > 0;
      const hasSecondary = Array.isArray(entry.secondaryKeys) && entry.secondaryKeys.length > 0;
      const hasTitle = String(entry.title || "").trim().length > 0;
      const hasContent = String(entry.content || "").trim().length > 0;
      return hasContent || hasKeys || hasSecondary || hasTitle || entry.constant;
    });
}

function normalizeLorebookRecord(item = {}) {
  const fromData = item?.data && typeof item.data === "object" ? item.data : {};
  const source = {
    ...(fromData && typeof fromData === "object" ? fromData : {}),
    ...(item && typeof item === "object" ? item : {}),
  };

  if (Object.prototype.hasOwnProperty.call(item || {}, "entries")) {
    source.entries = item.entries;
  }
  if (Object.prototype.hasOwnProperty.call(item || {}, "world_info")) {
    source.world_info = item.world_info;
  }

  const entries = parseLorebookEntries(source);
  const scanDepthRaw = Number(source.scanDepth ?? source.scan_depth ?? source.depth ?? defaults.lorebookSettings.scanDepth);
  const tokenBudgetRaw = Number(source.tokenBudget ?? source.token_budget ?? source.budget ?? defaults.lorebookSettings.tokenBudget);
  const recursionDepthRaw = Number(source.maxRecursionDepth ?? source.max_recursion_depth ?? source.recursionDepth ?? source.recursion_depth ?? 3);

  const defaultsSource = source.defaultEntrySettings && typeof source.defaultEntrySettings === "object"
    ? source.defaultEntrySettings
    : source.default_entry_settings && typeof source.default_entry_settings === "object"
      ? source.default_entry_settings
      : {};

  const nowIso = new Date().toISOString();

  const colorSource = String(source.color || source.coverColor || source.cover_color || LOREBOOK_V2_COLORS[0]);
  const color = LOREBOOK_V2_COLORS.includes(colorSource) ? colorSource : LOREBOOK_V2_COLORS[0];

  return {
    id: item.id || source.id || crypto.randomUUID(),
    name: String(source.name || source.title || item.name || "Imported lorebook"),
    description: String(source.description || item.description || ""),
    color,
    scanDepth: Math.min(64, Math.max(1, Number.isFinite(scanDepthRaw) ? scanDepthRaw : defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.isFinite(tokenBudgetRaw) ? tokenBudgetRaw : defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: Boolean(source.recursiveScanning ?? source.recursive_scanning ?? source.recursive ?? defaults.lorebookSettings.recursiveScanning),
    maxRecursionDepth: Math.min(12, Math.max(0, Number.isFinite(recursionDepthRaw) ? recursionDepthRaw : 3)),
    enabled: item.enabled !== false,
    attachmentScope: String(source.attachmentScope || source.attachment_scope || "specific"),
    attachmentCharacterId: String(source.attachmentCharacterId || source.attachment_character_id || ""),
    attachedChatIds: normalizeStringArray(source.attachedChatIds || source.attached_chat_ids || []),
    defaultEntrySettings: {
      insertionPosition: normalizeInsertionPosition(defaultsSource.insertionPosition || defaultsSource.insertion_position || "after_system"),
      scanDepth: Math.max(0, Number.parseInt(defaultsSource.scanDepth ?? defaultsSource.scan_depth ?? 10, 10) || 10),
      insertionOrder: Math.max(0, Number.parseInt(defaultsSource.insertionOrder ?? defaultsSource.insertion_order ?? 100, 10) || 100),
      enabled: defaultsSource.enabled !== false,
      caseSensitive: Boolean(defaultsSource.caseSensitive ?? defaultsSource.case_sensitive ?? false),
    },
    createdAt: String(source.createdAt || source.created_at || item.createdAt || nowIso),
    updatedAt: String(source.updatedAt || source.updated_at || item.updatedAt || nowIso),
    entries,
    data: {
      ...(source && typeof source === "object" ? source : {}),
      entries,
    },
  };
}

function matchesKeyword(text, keyword, caseSensitive = false) {
  const source = String(text || "");
  const needle = String(keyword || "").trim();
  if (!needle) return false;
  const flags = caseSensitive ? "g" : "gi";
  const pattern = `(^|[^A-Za-z0-9_])${escapeRegex(needle)}(?=$|[^A-Za-z0-9_])`;
  return new RegExp(pattern, flags).test(source);
}

function resolveLorebookPromptContext({ messages = [], character = null } = {}) {
  const chat = getCurrentChat();
  const attachedIds = new Set(Array.isArray(chat?.lorebookIds) ? chat.lorebookIds : []);
  const enabledLorebooks = state.lorebooks.filter((book) => {
    if (book?.enabled === false) return false;
    if (attachedIds.has(book.id)) return true;
    return lbv2IsLorebookAttachedToChat(book, chat);
  });
  if (!enabledLorebooks.length) return { beforeText: "", afterText: "", entries: [] };

  const activeSettings = getActiveChatLorebookSettings();
  const effectiveScanDepth = enabledLorebooks.length
    ? Math.min(...enabledLorebooks.map((book) => Math.max(1, Number.parseInt(book.scanDepth ?? activeSettings.scanDepth, 10) || activeSettings.scanDepth)))
    : activeSettings.scanDepth;
  const effectiveTokenBudget = enabledLorebooks.length
    ? Math.min(...enabledLorebooks.map((book) => Math.max(64, Number.parseInt(book.tokenBudget ?? activeSettings.tokenBudget, 10) || activeSettings.tokenBudget)))
    : activeSettings.tokenBudget;
  const recursiveEnabled = enabledLorebooks.some((book) => book.recursiveScanning !== false) && activeSettings.recursiveScanning;
  const maxRecursionDepth = Math.max(0, ...enabledLorebooks.map((book) => Number.parseInt(book?.maxRecursionDepth ?? 3, 10) || 3));

  const history = Array.isArray(messages)
    ? messages.filter((msg) => msg && (msg.role === "user" || msg.role === "assistant")).slice(-effectiveScanDepth)
    : [];
  const characterGreeting = String(character?.greeting || character?.data?.first_mes || "").trim();
  const characterSummary = [
    String(character?.name || "").trim(),
    String(character?.scenario || character?.data?.scenario || "").trim(),
    String(character?.personality || character?.data?.personality || character?.data?.description || "").trim(),
    characterGreeting,
  ].filter(Boolean).join("\n");
  let searchableText = [
    characterSummary,
    history.map((msg) => String(msg.content || "")).join("\n"),
  ].filter(Boolean).join("\n");

  const selectedEntries = [];
  const selectedIds = new Set();

  const shouldActivateEntry = (entry, searchText) => {
    if (!entry || entry.enabled === false) return false;
    if (entry.constant || entry.strategy === "constant") return true;

    if (!entry.keys.length && entry.strategy === "normal") return false;

    const matcher = (text, key) => {
      if (entry.useRegex) {
        try {
          const regex = new RegExp(String(key || ""), entry.caseSensitive ? "g" : "gi");
          return regex.test(String(text || ""));
        } catch {
          return false;
        }
      }

      if (entry.matchWholeWords) {
        return matchesKeyword(text, key, entry.caseSensitive);
      }

      const source = String(text || "");
      const needle = String(key || "").trim();
      if (!needle) return false;
      if (entry.caseSensitive) return source.includes(needle);
      return source.toLowerCase().includes(needle.toLowerCase());
    };

    const primaryMatchesCount = entry.keys.filter((key) => matcher(searchText, key)).length;
    const secondaryMatchesCount = entry.secondaryKeys.filter((key) => matcher(searchText, key)).length;

    const anyPrimary = primaryMatchesCount > 0;
    const allPrimary = entry.keys.length > 0 && primaryMatchesCount === entry.keys.length;
    const anySecondary = secondaryMatchesCount > 0;

    const logic = String(entry.activationLogic || "ANY").toUpperCase();
    if (logic === "ALL") return allPrimary;
    if (logic === "AND_ANY_SECONDARY") return anyPrimary && anySecondary;
    if (logic === "NOT") return !anyPrimary;

    if (!entry.selective || !entry.secondaryKeys.length) {
      return anyPrimary;
    }

    if (entry.selectiveLogic === "NOT") return !anySecondary && anyPrimary;
    if (entry.selectiveLogic === "OR") return anyPrimary || anySecondary;
    return anyPrimary && anySecondary;
  };

  const runPass = ({ recursivePass = false } = {}) => {
    let added = false;
    for (const lorebook of enabledLorebooks) {
      const entries = Array.isArray(lorebook.entries) ? lorebook.entries : [];
      for (const entry of entries) {
        if (selectedIds.has(entry.id)) continue;
        if (recursivePass && entry.excludeFromRecursion) continue;
        if (!shouldActivateEntry(entry, searchableText)) continue;

        const chanceRoll = Math.random() * 100;
        if (entry.probability < 100 && chanceRoll > entry.probability) continue;

        selectedIds.add(entry.id);
        selectedEntries.push({
          ...entry,
          _lorebookId: lorebook.id,
          _lorebookName: lorebook.name,
          _effectiveTokenBudget: Math.min(effectiveTokenBudget, lorebook.tokenBudget || effectiveTokenBudget),
        });
        if (!entry.preventRecursion) {
          searchableText += `\n${entry.content}`;
        }
        added = true;
      }
    }
    return added;
  };

  runPass({ recursivePass: false });
  if (recursiveEnabled) {
    const recursionCap = Math.min(12, Math.max(0, maxRecursionDepth));
    for (let pass = 0; pass < recursionCap; pass += 1) {
      if (!runPass({ recursivePass: true })) break;
    }
  }

  const sorted = selectedEntries
    .slice()
    .sort((a, b) => {
      const byPriority = Number(b.priority || 0) - Number(a.priority || 0);
      if (byPriority !== 0) return byPriority;
      const byOrder = Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0);
      if (byOrder !== 0) return byOrder;
      return String(a.title || "").localeCompare(String(b.title || ""));
    });

  const kept = [];
  let usedTokens = 0;
  const globalBudget = effectiveTokenBudget;
  for (const entry of sorted) {
    const text = applyMacrosToText(String(entry.content || "").trim(), character);
    if (!text) continue;
    const estimated = estimateTokenCount(text);

    const perEntryBudget = Number(entry.tokenBudget || 0) > 0
      ? Number(entry.tokenBudget)
      : globalBudget;

    if (estimated > perEntryBudget) continue;
    if (usedTokens + estimated > globalBudget) continue;

    kept.push({ ...entry, _rendered: text, _estimatedTokens: estimated });
    usedTokens += estimated;
  }

  kept.sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0));

  const groupedByName = new Map();
  const finalEntries = [];

  for (const entry of kept) {
    const groupName = String(entry.group || "").trim();
    if (!groupName) {
      finalEntries.push(entry);
      continue;
    }

    if (!groupedByName.has(groupName)) groupedByName.set(groupName, []);
    groupedByName.get(groupName).push(entry);
  }

  groupedByName.forEach((items) => {
    if (!items.length) return;
    if (items.length === 1) {
      finalEntries.push(items[0]);
      return;
    }

    const totalWeight = items.reduce((sum, item) => sum + Math.max(1, Number(item.groupWeight || 100)), 0);
    let roll = Math.random() * totalWeight;
    let picked = items[0];
    for (const item of items) {
      roll -= Math.max(1, Number(item.groupWeight || 100));
      if (roll <= 0) {
        picked = item;
        break;
      }
    }
    finalEntries.push(picked);
  });

  finalEntries.sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0));

  const beforeEntries = [];
  const afterEntries = [];

  finalEntries.forEach((entry) => {
    const position = String(entry.insertionPosition || "after_system");
    if (["after_system", "after_author_note", "bottom_chat"].includes(position)) {
      afterEntries.push(entry);
    } else {
      beforeEntries.push(entry);
    }
  });

  const beforeText = beforeEntries.map((entry) => entry._rendered).join("\n\n").trim();
  const afterText = afterEntries.map((entry) => entry._rendered).join("\n\n").trim();

  return {
    beforeText,
    afterText,
    entries: finalEntries,
    estimatedTokens: finalEntries.reduce((sum, entry) => sum + Number(entry._estimatedTokens || 0), 0),
    scanDepthUsed: effectiveScanDepth,
    tokenBudgetUsed: effectiveTokenBudget,
  };
}

function syncLorebookSettingsInputs() {
  if (els.lorebookScanDepthInput) {
    els.lorebookScanDepthInput.value = String(state.lorebookSettings.scanDepth || defaults.lorebookSettings.scanDepth);
  }
  if (els.lorebookTokenBudgetInput) {
    els.lorebookTokenBudgetInput.value = String(state.lorebookSettings.tokenBudget || defaults.lorebookSettings.tokenBudget);
  }
  if (els.lorebookRecursiveInput) {
    els.lorebookRecursiveInput.checked = state.lorebookSettings.recursiveScanning !== false;
  }
  setInlineStatus(els.lorebookSettingsStatus, "");
}

function getLabeledLorebookSettings(book = null) {
  const fallback = {
    scanDepth: Math.min(64, Math.max(1, Number.parseInt(state.lorebookSettings?.scanDepth, 10) || defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(state.lorebookSettings?.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: state.lorebookSettings?.recursiveScanning !== false,
  };

  if (!book || typeof book !== "object") return fallback;

  return {
    scanDepth: Math.min(64, Math.max(1, Number.parseInt(book.scanDepth, 10) || fallback.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(book.tokenBudget, 10) || fallback.tokenBudget)),
    recursiveScanning: book.recursiveScanning ?? fallback.recursiveScanning,
  };
}

function saveLorebookSettingsFromInputs() {
  const scanDepth = Math.min(64, Math.max(1, Number.parseInt(els.lorebookScanDepthInput?.value || "", 10) || defaults.lorebookSettings.scanDepth));
  const tokenBudget = Math.min(8000, Math.max(64, Number.parseInt(els.lorebookTokenBudgetInput?.value || "", 10) || defaults.lorebookSettings.tokenBudget));

  state.lorebookSettings.scanDepth = scanDepth;
  state.lorebookSettings.tokenBudget = tokenBudget;
  state.lorebookSettings.recursiveScanning = Boolean(els.lorebookRecursiveInput?.checked ?? true);

  persistState();
  syncLorebookSettingsInputs();
  setInlineStatus(els.lorebookSettingsStatus, "Lore settings saved");
  showToast("Lore settings saved", "success");
}

function syncLibraryLorebookFullscreenMode() {
  if (!els.libraryView) return;
  const lorebookPanelActive = Boolean(els.lbv2Panel)
    && state.activeView === "libraryView"
    && els.lbv2Panel.classList.contains("active")
    && !els.lbv2Panel.classList.contains("hidden");

  els.libraryView.classList.toggle("library-lorebook-open", lorebookPanelActive);
}

function setActiveLibraryPanel(panelId = "libraryHubPanel") {
  const target = panelId || "libraryHubPanel";
  const showHub = target === "libraryHubPanel";

  els.libraryPanels?.forEach((panel) => {
    panel.classList.toggle("hidden", panel.id !== target);
    panel.classList.toggle("active", panel.id === target);
  });

  if (els.libraryHubPanel) {
    els.libraryHubPanel.classList.toggle("hidden", !showHub);
    els.libraryHubPanel.classList.toggle("active", showHub);
  }

  els.libraryCategoryChips?.forEach((chip) => {
    chip.classList.toggle("active", chip.dataset.libraryPanel === target);
  });

  if (showHub) {
    renderLibraryHub();
  }

  if (!showHub && target === "libraryLorebooksPanel" && LOREBOOK_V2_ENABLED) {
    if (!lbv2State.activeLorebookId) {
      lbv2SetLevel(1);
    }
    lbv2State.activeEntryId = "";
    lbv2State.bulkMode = false;
    lbv2State.bulkSelectedEntryIds = new Set();
    renderLorebooksV2();
  }

  syncLibraryLorebookFullscreenMode();
}

function renderLibraryHub() {
  const chips = els.libraryCategoryChips || [];
  const query = String(libraryHubSearchQuery || "").trim().toLowerCase();
  let visibleCount = 0;

  chips.forEach((chip) => {
    const label = String(chip.textContent || "").toLowerCase();
    const keywords = String(chip.dataset.librarySearch || "").toLowerCase();
    const matches = !query || label.includes(query) || keywords.includes(query);
    chip.classList.toggle("hidden", !matches);
    if (matches) visibleCount += 1;
  });

  if (els.libraryHubEmptyState) {
    els.libraryHubEmptyState.classList.toggle("hidden", visibleCount > 0);
  }
}

function parseCommaSeparatedList(value = "") {
  return String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function formatDateShort(dateLike) {
  const date = new Date(dateLike || Date.now());
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

function formatRelativeTime(dateLike) {
  const date = new Date(dateLike || Date.now());
  if (Number.isNaN(date.getTime())) return "just now";
  const deltaMs = Date.now() - date.getTime();
  const abs = Math.abs(deltaMs);
  const minute = 60 * 1000;
  const hour = 60 * minute;
  const day = 24 * hour;

  if (abs < minute) return "just now";
  if (abs < hour) {
    const n = Math.max(1, Math.round(abs / minute));
    return `${n} min${n === 1 ? "" : "s"} ago`;
  }
  if (abs < day) {
    const n = Math.max(1, Math.round(abs / hour));
    return `${n} hour${n === 1 ? "" : "s"} ago`;
  }
  const n = Math.max(1, Math.round(abs / day));
  return `${n} day${n === 1 ? "" : "s"} ago`;
}

function countLorebookTokens(book) {
  const entries = Array.isArray(book?.entries) ? book.entries : [];
  return entries.reduce((sum, entry) => sum + estimateTokenCount(String(entry?.content || "")), 0);
}

function withLorebookTimestamps(book, fallback = {}) {
  const nowIso = new Date().toISOString();
  return {
    ...book,
    createdAt: String(book?.createdAt || book?.created_at || fallback.createdAt || nowIso),
    updatedAt: String(book?.updatedAt || book?.updated_at || fallback.updatedAt || nowIso),
  };
}

function withEntryTimestamps(entry, fallback = {}) {
  const nowIso = new Date().toISOString();
  return {
    ...entry,
    createdAt: String(entry?.createdAt || entry?.created_at || fallback.createdAt || nowIso),
    updatedAt: String(entry?.updatedAt || entry?.updated_at || fallback.updatedAt || nowIso),
  };
}

function openLorebookEntryEditor(index = null) {
  if (!els.lorebookEntryEditorCard) return;

  const hasIndex = Number.isInteger(index) && index >= 0 && lorebookEntryDraft[index];
  const entry = hasIndex ? lorebookEntryDraft[index] : {
    title: "",
    keys: [],
    secondaryKeys: [],
    content: "",
    priority: 100,
    probability: 100,
    caseSensitive: false,
    selective: false,
    selectiveLogic: "AND",
    constant: false,
    enabled: true,
  };

  els.lorebookEntryEditorCard.classList.remove("hidden");
  if (hasIndex) {
    els.lorebookEntryEditorCard.dataset.editingIndex = String(index);
  } else {
    delete els.lorebookEntryEditorCard.dataset.editingIndex;
  }

  if (els.lorebookEntryTitleInput) els.lorebookEntryTitleInput.value = entry.title || "";
  if (els.lorebookEntryKeysInput) els.lorebookEntryKeysInput.value = Array.isArray(entry.keys) ? entry.keys.join(", ") : "";
  if (els.lorebookEntrySecondaryKeysInput) els.lorebookEntrySecondaryKeysInput.value = Array.isArray(entry.secondaryKeys) ? entry.secondaryKeys.join(", ") : "";
  if (els.lorebookEntryContentInput) els.lorebookEntryContentInput.value = entry.content || entry.entry || "";
  if (els.lorebookEntryPriorityInput) els.lorebookEntryPriorityInput.value = String(Number.isFinite(Number(entry.priority)) ? Number(entry.priority) : 100);
  if (els.lorebookEntryProbabilityInput) els.lorebookEntryProbabilityInput.value = String(Number.isFinite(Number(entry.probability)) ? Number(entry.probability) : 100);
  if (els.lorebookEntryEnabledInput) els.lorebookEntryEnabledInput.checked = entry.enabled !== false;
  if (els.lorebookEntryConstantInput) els.lorebookEntryConstantInput.checked = Boolean(entry.constant);
  if (els.lorebookEntryCaseSensitiveInput) els.lorebookEntryCaseSensitiveInput.checked = Boolean(entry.caseSensitive);
  if (els.lorebookEntrySelectiveInput) els.lorebookEntrySelectiveInput.checked = Boolean(entry.selective);
  if (els.lorebookEntrySelectiveLogicInput) els.lorebookEntrySelectiveLogicInput.value = entry.selectiveLogic || "AND";

  if (els.lorebookEntryDeleteBtn) {
    els.lorebookEntryDeleteBtn.classList.toggle("hidden", !hasIndex);
  }

  els.lorebookEntryTitleInput?.focus();
}

function closeLorebookEntryEditor() {
  if (!els.lorebookEntryEditorCard) return;
  els.lorebookEntryEditorCard.classList.add("hidden");
  delete els.lorebookEntryEditorCard.dataset.editingIndex;
  if (els.lorebookEntryTitleInput) els.lorebookEntryTitleInput.value = "";
  if (els.lorebookEntryKeysInput) els.lorebookEntryKeysInput.value = "";
  if (els.lorebookEntrySecondaryKeysInput) els.lorebookEntrySecondaryKeysInput.value = "";
  if (els.lorebookEntryContentInput) els.lorebookEntryContentInput.value = "";
  if (els.lorebookEntryPriorityInput) els.lorebookEntryPriorityInput.value = "100";
  if (els.lorebookEntryProbabilityInput) els.lorebookEntryProbabilityInput.value = "100";
  if (els.lorebookEntryEnabledInput) els.lorebookEntryEnabledInput.checked = true;
  if (els.lorebookEntryConstantInput) els.lorebookEntryConstantInput.checked = false;
  if (els.lorebookEntryCaseSensitiveInput) els.lorebookEntryCaseSensitiveInput.checked = false;
  if (els.lorebookEntrySelectiveInput) els.lorebookEntrySelectiveInput.checked = false;
  if (els.lorebookEntrySelectiveLogicInput) els.lorebookEntrySelectiveLogicInput.value = "AND";
}

function saveLorebookEntryFromEditor(options = {}) {
  const skipClose = Boolean(options?.skipClose);
  const editingIndex = Number.parseInt(els.lorebookEntryEditorCard?.dataset?.editingIndex || "-1", 10);
  const fallbackTitle = Number.isFinite(editingIndex) && editingIndex >= 0
    ? `Entry ${editingIndex + 1}`
    : `Entry ${lorebookEntryDraft.length + 1}`;
  const title = els.lorebookEntryTitleInput?.value?.trim() || fallbackTitle;
  const keys = parseCommaSeparatedList(els.lorebookEntryKeysInput?.value || "");
  const secondaryKeys = parseCommaSeparatedList(els.lorebookEntrySecondaryKeysInput?.value || "");
  const content = els.lorebookEntryContentInput?.value?.trim() || "";

  const payload = normalizeLorebookEntry({
    id: crypto.randomUUID(),
    title,
    keys,
    secondary_keys: secondaryKeys,
    content,
    priority: Number.parseInt(els.lorebookEntryPriorityInput?.value || "100", 10),
    probability: Number.parseInt(els.lorebookEntryProbabilityInput?.value || "100", 10),
    case_sensitive: Boolean(els.lorebookEntryCaseSensitiveInput?.checked),
    selective: Boolean(els.lorebookEntrySelectiveInput?.checked),
    selective_logic: els.lorebookEntrySelectiveLogicInput?.value || "AND",
    constant: Boolean(els.lorebookEntryConstantInput?.checked),
    enabled: Boolean(els.lorebookEntryEnabledInput?.checked ?? true),
  }, 0);

  if (Number.isFinite(editingIndex) && editingIndex >= 0 && lorebookEntryDraft[editingIndex]) {
    lorebookEntryDraft[editingIndex] = {
      ...lorebookEntryDraft[editingIndex],
      ...payload,
      id: lorebookEntryDraft[editingIndex].id || payload.id,
    };
    setInlineStatus(els.lorebookEditorStatus, `Updated entry ${editingIndex + 1}`);
  } else {
    lorebookEntryDraft.push(payload);
    setInlineStatus(els.lorebookEditorStatus, "Entry added");
  }

  lorebookEntryDraft = lorebookEntryDraft.map((entry, index) => ({ ...entry, insertionOrder: index }));
  lorebookEntrySearchQuery = "";
  if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
  if (!skipClose) {
    closeLorebookEntryEditor();
  }
  renderLorebookEntryDraft();
  return true;
}

function deleteLorebookEntryFromEditor() {
  const editingIndex = Number.parseInt(els.lorebookEntryEditorCard?.dataset?.editingIndex || "-1", 10);
  if (!Number.isFinite(editingIndex) || editingIndex < 0 || !lorebookEntryDraft[editingIndex]) {
    closeLorebookEntryEditor();
    return;
  }

  lorebookEntryDraft.splice(editingIndex, 1);
  lorebookEntryDraft = lorebookEntryDraft.map((entry, index) => ({ ...entry, insertionOrder: index }));
  closeLorebookEntryEditor();
  renderLorebookEntryDraft();
  setInlineStatus(els.lorebookEditorStatus, "Entry deleted");
}

function showLorebookForm(lorebook = null) {
  const sourceLorebook = lorebook?.id ? (state.lorebooks.find((book) => book.id === lorebook.id) || lorebook) : lorebook;
  const normalized = sourceLorebook ? normalizeLorebookRecord(sourceLorebook) : normalizeLorebookRecord({ name: "", description: "", entries: [] });
  if (els.lorebookEditorView) {
    els.lorebookEditorView.dataset.editingId = normalized.id || lorebook?.id || "";
  }

  if (els.lorebookNameInput) els.lorebookNameInput.value = normalized.name || "";
  if (els.lorebookDescriptionInput) els.lorebookDescriptionInput.value = normalized.description || "";
  lorebookEntryDraft = (Array.isArray(normalized.entries) ? normalized.entries : [])
    .slice()
    .sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0))
    .map((entry, index) => ({ ...entry, insertionOrder: index }));

  lorebookEntrySearchQuery = "";
  if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
  closeLorebookEntryEditor();
  renderLorebookEntryDraft();
  setInlineStatus(els.lorebookEditorStatus, "");
  switchView("lorebookEditorView");
  els.lorebookNameInput?.focus();
}

function hideLorebookForm() {
  if (els.lorebookEditorView) {
    delete els.lorebookEditorView.dataset.editingId;
  }
  if (els.lorebookNameInput) els.lorebookNameInput.value = "";
  if (els.lorebookDescriptionInput) els.lorebookDescriptionInput.value = "";
  lorebookEntryDraft = [];
  lorebookEntrySearchQuery = "";
  if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
  closeLorebookEntryEditor();
  renderLorebookEntryDraft();
  setInlineStatus(els.lorebookEditorStatus, "");
  switchView("libraryView");
  setActiveLibraryPanel("libraryHubPanel");
  renderLibraryHub();
}

function clearLorebookEntryLongPressState() {
  if (lorebookEntryLongPress.timer) {
    clearTimeout(lorebookEntryLongPress.timer);
  }
  if (lorebookEntryLongPress.moveHandler) {
    document.removeEventListener("pointermove", lorebookEntryLongPress.moveHandler);
  }
  if (lorebookEntryLongPress.upHandler) {
    document.removeEventListener("pointerup", lorebookEntryLongPress.upHandler);
    document.removeEventListener("pointercancel", lorebookEntryLongPress.upHandler);
  }
  if (lorebookEntryLongPress.autoScrollRaf) {
    cancelAnimationFrame(lorebookEntryLongPress.autoScrollRaf);
  }

  els.lorebookEntryList?.classList.remove("drag-mode");
  els.lorebookEntryList?.querySelectorAll(".lorebook-entry-item").forEach((node) => {
    node.classList.remove("drag-source", "drop-target", "longpress-arming", "longpress-active");
    node.style.transform = "";
  });

  lorebookEntryLongPress = {
    timer: null,
    pointerId: null,
    sourceIndex: -1,
    targetIndex: -1,
    active: false,
    moveHandler: null,
    upHandler: null,
    autoScrollRaf: null,
    latestClientX: 0,
    latestClientY: 0,
    scrollContainer: null,
  };
}

function isScrollableY(element) {
  if (!element || !(element instanceof Element)) return false;
  const style = getComputedStyle(element);
  const allowsScroll = /(auto|scroll|overlay)/.test(style.overflowY || "");
  return allowsScroll && element.scrollHeight > element.clientHeight + 2;
}

function getDragScrollContainer(startElement) {
  let node = startElement;
  while (node && node !== document.body) {
    if (isScrollableY(node)) return node;
    node = node.parentElement;
  }
  return document.scrollingElement || document.documentElement;
}

function getScrollContainerRect(container) {
  if (!container) {
    return { top: 0, bottom: window.innerHeight };
  }
  if (container === document.scrollingElement || container === document.documentElement || container === document.body) {
    return { top: 0, bottom: window.innerHeight };
  }
  return container.getBoundingClientRect();
}

function applyDragAutoScroll(container, clientY) {
  if (!container) return;

  const rect = getScrollContainerRect(container);
  const edgeSize = 72;
  const maxStep = 18;
  let delta = 0;

  if (clientY < rect.top + edgeSize) {
    const t = Math.max(0, (rect.top + edgeSize - clientY) / edgeSize);
    delta = -Math.ceil(maxStep * t);
  } else if (clientY > rect.bottom - edgeSize) {
    const t = Math.max(0, (clientY - (rect.bottom - edgeSize)) / edgeSize);
    delta = Math.ceil(maxStep * t);
  }

  if (!delta) return;

  if (container === document.scrollingElement || container === document.documentElement || container === document.body) {
    window.scrollBy(0, delta);
    return;
  }

  container.scrollTop += delta;
}

function renderLorebookEntryDraft() {
  if (!els.lorebookEntryList) return;

  clearLorebookEntryLongPressState();
  els.lorebookEntryList.innerHTML = "";

  const query = lorebookEntrySearchQuery.trim().toLowerCase();
  const visibleEntries = lorebookEntryDraft
    .map((entry, index) => ({ entry, index }))
    .filter(({ entry }) => {
      if (!query) return true;
      const haystack = `${entry.title || ""} ${(entry.keys || []).join(" ")} ${(entry.secondaryKeys || []).join(" ")} ${entry.content || ""}`.toLowerCase();
      return haystack.includes(query);
    });

  if (!visibleEntries.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">${query ? "No entries match your search." : "No entries yet. Create one."}</div>`;
    els.lorebookEntryList.appendChild(empty);
    return;
  }

  visibleEntries.forEach(({ entry, index }) => {
    const card = document.createElement("article");
    card.className = "entity-card lorebook-entry-item";
    card.dataset.entryIndex = String(index);

    const keyPreview = Array.isArray(entry.keys) ? entry.keys.slice(0, 4).join(", ") : "";
    const contentText = String(entry.content || entry.entry || "").trim();
    const contentPreview = contentText.slice(0, 260);

    card.innerHTML = `
      <div class="entity-card-header">
        <h3>Entry ${index + 1}: ${escapeHtml(entry.title || "Untitled")}</h3>
        <span class="entity-meta">${(Array.isArray(entry.keys) ? entry.keys.length : 0)} keys</span>
      </div>
      <div class="entity-meta">${escapeHtml(keyPreview || "No keys")}</div>
      <div class="entity-meta">${escapeHtml(contentPreview || "No content")}</div>
      <div class="entity-card-actions">
        <button class="secondary" type="button" data-action="edit-entry">Edit</button>
      </div>
    `;

    card.querySelector('[data-action="edit-entry"]')?.addEventListener("click", (event) => {
      event.stopPropagation();
      openLorebookEntryEditor(index);
    });

    card.addEventListener("pointerdown", (event) => {
      if (event.target?.closest?.("button")) return;
      if (event.button !== undefined && event.button !== 0) return;

      clearLorebookEntryLongPressState();
      card.classList.add("longpress-arming");

      lorebookEntryLongPress.pointerId = event.pointerId;
      lorebookEntryLongPress.sourceIndex = index;
      lorebookEntryLongPress.targetIndex = index;
      lorebookEntryLongPress.latestClientX = event.clientX;
      lorebookEntryLongPress.latestClientY = event.clientY;
      lorebookEntryLongPress.scrollContainer = getDragScrollContainer(els.lorebookEntryList || card);

      const originX = event.clientX;
      const originY = event.clientY;

      const removePreLongPressListeners = () => {
        document.removeEventListener("pointermove", preLongPressMove);
        document.removeEventListener("pointerup", preLongPressEnd);
        document.removeEventListener("pointercancel", preLongPressEnd);
      };

      const preLongPressMove = (moveEvent) => {
        if (moveEvent.pointerId !== lorebookEntryLongPress.pointerId) return;
        lorebookEntryLongPress.latestClientX = moveEvent.clientX;
        lorebookEntryLongPress.latestClientY = moveEvent.clientY;

        // User is scrolling, cancel long-press arming and let normal scroll continue.
        const dx = Math.abs(moveEvent.clientX - originX);
        const dy = Math.abs(moveEvent.clientY - originY);
        if (!lorebookEntryLongPress.active && (dy > 10 || dx > 10)) {
          if (lorebookEntryLongPress.timer) clearTimeout(lorebookEntryLongPress.timer);
          lorebookEntryLongPress.timer = null;
          card.classList.remove("longpress-arming");
          removePreLongPressListeners();
        }
      };

      const preLongPressEnd = (upEvent) => {
        if (upEvent.pointerId !== lorebookEntryLongPress.pointerId) return;
        removePreLongPressListeners();
        if (!lorebookEntryLongPress.active) {
          clearLorebookEntryLongPressState();
        }
      };

      lorebookEntryLongPress.timer = setTimeout(() => {
        lorebookEntryLongPress.active = true;
        card.classList.remove("longpress-arming");
        card.classList.add("drag-source", "longpress-active");
        els.lorebookEntryList?.classList.add("drag-mode");

        lorebookEntryLongPress.moveHandler = (moveEvent) => {
          if (moveEvent.pointerId !== lorebookEntryLongPress.pointerId) return;

          lorebookEntryLongPress.latestClientX = moveEvent.clientX;
          lorebookEntryLongPress.latestClientY = moveEvent.clientY;

          // prevent page pan only after drag mode starts
          moveEvent.preventDefault();

          const dy = moveEvent.clientY - originY;
          card.style.transform = `translateY(${dy}px) scale(1.02)`;

          applyDragAutoScroll(lorebookEntryLongPress.scrollContainer, lorebookEntryLongPress.latestClientY);

          const hovered = document.elementFromPoint(moveEvent.clientX, moveEvent.clientY)?.closest?.(".lorebook-entry-item");
          if (!hovered || hovered.parentElement !== els.lorebookEntryList) return;

          const hoverIndex = Number.parseInt(hovered.dataset.entryIndex || "-1", 10);
          if (!Number.isFinite(hoverIndex) || hoverIndex < 0) return;

          lorebookEntryLongPress.targetIndex = hoverIndex;
          els.lorebookEntryList?.querySelectorAll(".lorebook-entry-item").forEach((itemNode) => itemNode.classList.remove("drop-target"));
          if (hoverIndex !== lorebookEntryLongPress.sourceIndex) {
            hovered.classList.add("drop-target");
          }
        };

        lorebookEntryLongPress.upHandler = (upEvent) => {
          if (upEvent.pointerId !== lorebookEntryLongPress.pointerId) return;

          const from = lorebookEntryLongPress.sourceIndex;
          const to = lorebookEntryLongPress.targetIndex;

          if (
            lorebookEntryLongPress.active
            && Number.isFinite(from)
            && Number.isFinite(to)
            && from >= 0
            && to >= 0
            && from !== to
            && lorebookEntryDraft[from]
          ) {
            const [movedEntry] = lorebookEntryDraft.splice(from, 1);
            const insertionIndex = from < to ? Math.max(0, to - 1) : to;
            lorebookEntryDraft.splice(insertionIndex, 0, movedEntry);
            lorebookEntryDraft = lorebookEntryDraft.map((item, itemIndex) => ({
              ...item,
              insertionOrder: itemIndex,
            }));
            renderLorebookEntryDraft();
            setInlineStatus(els.lorebookEditorStatus, `Moved entry ${from + 1} to ${insertionIndex + 1}`);
          } else {
            clearLorebookEntryLongPressState();
            renderLorebookEntryDraft();
          }
        };

        const autoScrollLoop = () => {
          if (!lorebookEntryLongPress.active) return;
          applyDragAutoScroll(lorebookEntryLongPress.scrollContainer, lorebookEntryLongPress.latestClientY);
          lorebookEntryLongPress.autoScrollRaf = requestAnimationFrame(autoScrollLoop);
        };

        document.addEventListener("pointermove", lorebookEntryLongPress.moveHandler, { passive: false });
        document.addEventListener("pointerup", lorebookEntryLongPress.upHandler);
        document.addEventListener("pointercancel", lorebookEntryLongPress.upHandler);
        lorebookEntryLongPress.autoScrollRaf = requestAnimationFrame(autoScrollLoop);
      }, 320);

      document.addEventListener("pointermove", preLongPressMove);
      document.addEventListener("pointerup", preLongPressEnd);
      document.addEventListener("pointercancel", preLongPressEnd);
    });

    card.addEventListener("contextmenu", (event) => event.preventDefault());
    els.lorebookEntryList.appendChild(card);
  });
}

function tryParseLooseJson(text = "") {
  const source = String(text || "").trim();
  if (!source) throw new Error("Empty file");

  try {
    return JSON.parse(source);
  } catch {
    // continue
  }

  // Remove dangling commas first.
  try {
    const noTrailingCommas = source.replace(/,\s*([}\]])/g, "$1");
    return JSON.parse(noTrailingCommas);
  } catch {
    // continue
  }

  // If the file is truncated, try to salvage by balancing braces/brackets.
  try {
    let candidate = source
      .replace(/,\s*([}\]])/g, "$1")
      .replace(/\u0000/g, "")
      .trim();

    let inString = false;
    let escape = false;
    let curly = 0;
    let square = 0;

    for (let i = 0; i < candidate.length; i += 1) {
      const ch = candidate[i];
      if (escape) {
        escape = false;
        continue;
      }
      if (ch === "\\") {
        escape = true;
        continue;
      }
      if (ch === '"') {
        inString = !inString;
        continue;
      }
      if (inString) continue;
      if (ch === "{") curly += 1;
      else if (ch === "}") curly -= 1;
      else if (ch === "[") square += 1;
      else if (ch === "]") square -= 1;
    }

    if (inString) {
      candidate += '"';
    }

    while (square > 0) {
      candidate += "]";
      square -= 1;
    }
    while (curly > 0) {
      candidate += "}";
      curly -= 1;
    }

    // Clean commas again after auto-closing brackets/braces.
    candidate = candidate.replace(/,\s*([}\]])/g, "$1");

    return JSON.parse(candidate);
  } catch {
    // continue
  }

  throw new Error("Could not parse JSON (possibly truncated or invalid)");
}

async function importLorebookIntoEditor(event) {
  const [file] = event?.target?.files || [];
  if (!file) return;

  try {
    const text = await file.text();
    const parsed = tryParseLooseJson(text);
    const normalized = normalizeLorebookRecord({
      ...parsed,
      name: parsed?.name || parsed?.title || file.name.replace(/\.json$/i, "") || "Imported lorebook",
      description: parsed?.description || "",
    });

    if (els.lorebookEditorView) {
      els.lorebookEditorView.dataset.editingId = "";
    }
    if (els.lorebookNameInput) els.lorebookNameInput.value = normalized.name || "";
    if (els.lorebookDescriptionInput) els.lorebookDescriptionInput.value = normalized.description || "";

    lorebookEntryDraft = (Array.isArray(normalized.entries) ? normalized.entries : [])
      .slice()
      .sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0))
      .map((entry, index) => ({ ...entry, insertionOrder: index }));

    // Fallback parser for unsupported but common lorebook shapes.
    if (!lorebookEntryDraft.length) {
      const scanObjects = [];
      const queue = [parsed];
      const seen = new Set();
      while (queue.length) {
        const current = queue.shift();
        if (!current || typeof current !== "object") continue;
        if (seen.has(current)) continue;
        seen.add(current);

        if (
          current !== parsed
          && (
            Object.prototype.hasOwnProperty.call(current, "content")
            || Object.prototype.hasOwnProperty.call(current, "text")
            || Object.prototype.hasOwnProperty.call(current, "keys")
            || Object.prototype.hasOwnProperty.call(current, "key")
            || Object.prototype.hasOwnProperty.call(current, "keywords")
          )
        ) {
          scanObjects.push(current);
        }

        Object.values(current).forEach((value) => {
          if (value && typeof value === "object") queue.push(value);
        });
      }

      if (scanObjects.length) {
        lorebookEntryDraft = scanObjects
          .map((entry, index) => normalizeLorebookEntry(entry, index))
          .filter((entry) => {
            const hasKeys = Array.isArray(entry.keys) && entry.keys.length > 0;
            const hasSecondary = Array.isArray(entry.secondaryKeys) && entry.secondaryKeys.length > 0;
            const hasTitle = String(entry.title || "").trim().length > 0;
            const hasContent = String(entry.content || "").trim().length > 0;
            return hasContent || hasKeys || hasSecondary || hasTitle || entry.constant;
          })
          .map((entry, index) => ({ ...entry, insertionOrder: index }));
      }
    }

    lorebookEntrySearchQuery = "";
    if (els.lorebookEntrySearchInput) els.lorebookEntrySearchInput.value = "";
    closeLorebookEntryEditor();
    renderLorebookEntryDraft();
    setInlineStatus(els.lorebookEditorStatus, `Imported ${normalized.name} (${lorebookEntryDraft.length} entries)`);
    showToast("Lorebook loaded into editor", "success");
  } catch (error) {
    setInlineStatus(els.lorebookEditorStatus, `Could not import lorebook JSON: ${error?.message || "invalid format"}`);
    showToast("Lorebook import failed", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function saveLorebookFromForm() {
  const name = els.lorebookNameInput?.value?.trim() || "";
  const description = els.lorebookDescriptionInput?.value?.trim() || "";

  if (!name) {
    setStatus("Add a lorebook name");
    showToast("Lorebook save failed", "error");
    return;
  }

  if (els.lorebookEntryEditorCard && !els.lorebookEntryEditorCard.classList.contains("hidden")) {
    const ok = saveLorebookEntryFromEditor({ skipClose: true });
    if (!ok) {
      showToast("Fix entry fields before saving lorebook", "error");
      return;
    }
  }

  lorebookEntryDraft = (Array.isArray(lorebookEntryDraft) ? lorebookEntryDraft : [])
    .map((entry, index) => normalizeLorebookEntry({
      ...entry,
      insertion_order: index,
    }, index));

  const editingId = els.lorebookEditorView?.dataset?.editingId || "";
  const existing = editingId ? state.lorebooks.find((book) => book.id === editingId) : null;
  const nowIso = new Date().toISOString();

  const normalized = normalizeLorebookRecord({
    ...(existing || {}),
    id: editingId || existing?.id || crypto.randomUUID(),
    name,
    description,
    color: existing?.color || LOREBOOK_V2_COLORS[0],
    createdAt: existing?.createdAt || nowIso,
    updatedAt: nowIso,
    entries: lorebookEntryDraft.map((entry, index) => ({
      id: entry.id || crypto.randomUUID(),
      title: entry.title,
      comment: entry.comment || "",
      keys: Array.isArray(entry.keys) ? entry.keys : [],
      secondary_keys: Array.isArray(entry.secondaryKeys) ? entry.secondaryKeys : [],
      activation_logic: entry.activationLogic || "ANY",
      match_whole_words: Boolean(entry.matchWholeWords),
      use_regex: Boolean(entry.useRegex),
      content: String(entry.content || "").trim(),
      insertion_order: index,
      insertion_position: entry.insertionPosition || "after_system",
      insertion_depth: Number.isFinite(Number(entry.insertionDepth)) ? Number(entry.insertionDepth) : 4,
      role: entry.role || "system",
      priority: Number.isFinite(Number(entry.priority)) ? Number(entry.priority) : 100,
      case_sensitive: Boolean(entry.caseSensitive),
      selective: Boolean(entry.selective),
      selective_logic: entry.selectiveLogic || "AND",
      scan_depth: Number.isFinite(Number(entry.scanDepth)) ? Number(entry.scanDepth) : 0,
      token_budget: Number.isFinite(Number(entry.tokenBudget)) ? Number(entry.tokenBudget) : 0,
      constant: Boolean(entry.constant),
      strategy: entry.strategy || (entry.constant ? "constant" : "normal"),
      prevent_recursion: Boolean(entry.preventRecursion),
      exclude_from_recursion: Boolean(entry.excludeFromRecursion),
      probability: Number.isFinite(Number(entry.probability)) ? Number(entry.probability) : 100,
      group: entry.group || "",
      group_weight: Number.isFinite(Number(entry.groupWeight)) ? Number(entry.groupWeight) : 100,
      category: entry.category || "",
      tags: Array.isArray(entry.tags) ? entry.tags : [],
      automation_id: entry.automationId || "",
      linked_entries: Array.isArray(entry.linkedEntries) ? entry.linkedEntries : [],
      sticky: Number.isFinite(Number(entry.sticky)) ? Number(entry.sticky) : 0,
      cooldown: Number.isFinite(Number(entry.cooldown)) ? Number(entry.cooldown) : 0,
      delay: Number.isFinite(Number(entry.delay)) ? Number(entry.delay) : 0,
      turn_start: Number.isFinite(Number(entry.turnStart)) ? Number(entry.turnStart) : 0,
      turn_end: Number.isFinite(Number(entry.turnEnd)) ? Number(entry.turnEnd) : null,
      enabled: entry.enabled !== false,
      created_at: entry.createdAt || nowIso,
      updated_at: nowIso,
    })),
  });

  const index = editingId ? state.lorebooks.findIndex((book) => book.id === editingId) : -1;

  if (index >= 0) {
    state.lorebooks[index] = normalized;
    setStatus("Lorebook updated");
  } else {
    state.lorebooks.unshift(normalized);
    setStatus("Lorebook created");
  }

  persistState();
  renderLorebooks();
  renderDrawerLorebooks();
  showLorebookForm(normalized);
  setInlineStatus(els.lorebookEditorStatus, `Lorebook saved (${normalized.entries.length} entries)`);
  showToast("Lorebook saved", "success");
}

async function importLorebooks(event) {
  const files = Array.from(event?.target?.files || []);
  if (!files.length) return;

  const activeChat = getCurrentChat();

  try {
    for (const file of files) {
      const text = await file.text();
      const parsed = tryParseLooseJson(text);
      const items = Array.isArray(parsed?.lorebooks)
        ? parsed.lorebooks
        : Array.isArray(parsed)
          ? parsed
          : [parsed];

      items.forEach((item) => {
        const nowIso = new Date().toISOString();
        const normalized = normalizeLorebookRecord({
          ...item,
          name: item?.name || item?.title || file.name.replace(/\.json$/i, "") || "Imported lorebook",
          color: item?.color || LOREBOOK_V2_COLORS[0],
          createdAt: item?.createdAt || nowIso,
          updatedAt: nowIso,
          attachmentScope: item?.attachmentScope || "specific",
          attachedChatIds: activeChat?.id
            ? [activeChat.id]
            : normalizeStringArray(item?.attachedChatIds || item?.attached_chat_ids || []),
        });
        state.lorebooks.unshift(normalized);
      });
    }

    if (activeChat?.id) {
      const attachIds = (Array.isArray(state.lorebooks) ? state.lorebooks : [])
        .filter((book) => Array.isArray(book.attachedChatIds) && book.attachedChatIds.includes(activeChat.id))
        .map((book) => book.id);
      activeChat.lorebookIds = [...new Set([...(activeChat.lorebookIds || []), ...attachIds])];
      syncCurrentChatFromState();
    }

    persistState();
    renderLorebooks();
    renderDrawerLorebooks();
    setStatus("Lorebook imported");
    showToast("Lorebook imported", "success");
  } catch (error) {
    setStatus(`Could not import lorebook JSON: ${error?.message || "invalid format"}`);
    showToast("Lorebook import failed", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function exportLorebooks() {
  downloadJson("maitavern-lorebooks.json", state.lorebooks);
  setStatus("Lorebooks exported");
  showToast("Lorebooks exported", "success");
}

function exportLorebook(lorebookId) {
  const lorebook = state.lorebooks.find((item) => item.id === lorebookId);
  if (!lorebook) return;
  downloadJson(`${slugify(lorebook.name || "lorebook")}.json`, lorebook);
  setStatus("Lorebook exported");
  showToast("Lorebook exported", "success");
}

function removeLorebook(lorebookId) {
  state.lorebooks = state.lorebooks.filter((item) => item.id !== lorebookId);
  state.chats.forEach((chat) => {
    if (!Array.isArray(chat.lorebookIds)) return;
    chat.lorebookIds = chat.lorebookIds.filter((id) => id !== lorebookId);
  });

  if (lbv2State.activeLorebookId === lorebookId) {
    lbv2State.activeLorebookId = "";
    lbv2State.activeEntryId = "";
    lbv2State.bulkSelectedEntryIds = new Set();
    lbv2State.bulkMode = false;
  }

  persistState();
  renderLorebooks();
  renderDrawerLorebooks();
  setStatus("Lorebook removed");
  showToast("Lorebook deleted", "success");
}

function lbv2GetActiveLorebook() {
  if (!lbv2State.activeLorebookId) return null;
  return state.lorebooks.find((book) => book.id === lbv2State.activeLorebookId) || null;
}

function lbv2GetActiveEntry() {
  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook || !lbv2State.activeEntryId) return null;
  return (Array.isArray(lorebook.entries) ? lorebook.entries : []).find((entry) => entry.id === lbv2State.activeEntryId) || null;
}

function lbv2GetEntryById(lorebook, entryId) {
  if (!lorebook || !Array.isArray(lorebook.entries) || !entryId) return null;
  return lorebook.entries.find((entry) => entry.id === entryId) || null;
}

function lbv2IsLorebookAttachedToChat(book, chat = null) {
  if (!book) return false;
  const activeChat = chat || getCurrentChat();
  if (!activeChat?.id) return false;

  const scope = String(book.attachmentScope || "specific");
  if (scope === "global") return true;
  if (scope === "character") {
    return Boolean(activeChat.characterId && book.attachmentCharacterId && activeChat.characterId === book.attachmentCharacterId);
  }

  const attachedIds = Array.isArray(book.attachedChatIds) ? book.attachedChatIds : [];
  if (attachedIds.length) return attachedIds.includes(activeChat.id);

  return Array.isArray(activeChat.lorebookIds) && activeChat.lorebookIds.includes(book.id);
}

function lbv2ToggleLorebookAttachment(bookId) {
  const chat = getCurrentChat();
  if (!chat) {
    showToast("Open a chat first to attach lorebooks", "error");
    return;
  }

  const book = state.lorebooks.find((item) => item.id === bookId);
  if (!book) return;

  if (book.attachmentScope !== "specific") {
    const previouslyAttachedIds = (state.chats || [])
      .filter((item) => lbv2IsLorebookAttachedToChat(book, item))
      .map((item) => item.id);
    book.attachmentScope = "specific";
    book.attachedChatIds = [...new Set(previouslyAttachedIds)];
  }

  if (!Array.isArray(chat.lorebookIds)) chat.lorebookIds = [];
  if (!Array.isArray(book.attachedChatIds)) book.attachedChatIds = [];

  const set = new Set(chat.lorebookIds.map((id) => String(id || "")).filter(Boolean));
  const chatSet = new Set(book.attachedChatIds.map((id) => String(id || "")).filter(Boolean));

  if (set.has(bookId)) {
    set.delete(bookId);
    chatSet.delete(chat.id);
  } else {
    set.add(bookId);
    chatSet.add(chat.id);
  }

  chat.lorebookIds = [...set];
  book.attachedChatIds = [...chatSet];

  syncCurrentChatFromState();
  persistState();
  renderDrawerLorebooks();
  renderLorebooks();
}

function lbv2SetLevel(level = 1) {
  lbv2State.level = Math.min(3, Math.max(1, Number.parseInt(level, 10) || 1));
  lbv2SyncStack();
}

function lbv2NavigateTo(level = 1) {
  const nextLevel = Math.min(3, Math.max(1, Number.parseInt(level, 10) || 1));
  lbv2State.level = nextLevel;
  if (els.lbv2Stack) {
    els.lbv2Stack.dataset.level = String(nextLevel);
    els.lbv2Stack.dataset.anim = "slide";
  }
  lbv2SyncStack();
}

function lbv2NavigateBack() {
  if (lbv2State.level <= 1) {
    setActiveLibraryPanel("libraryHubPanel");
    return;
  }

  const target = lbv2State.level - 1;
  lbv2NavigateTo(target);
  if (target === 1) {
    lbv2State.activeEntryId = "";
    lbv2State.bulkMode = false;
    lbv2State.bulkSelectedEntryIds = new Set();
  }
  if (target === 2) {
    lbv2State.activeEntryId = "";
  }
}

function lbv2SyncStack() {
  const levels = [
    { node: els.lbv2Level1, level: 1 },
    { node: els.lbv2Level2, level: 2 },
    { node: els.lbv2Level3, level: 3 },
  ];

  levels.forEach(({ node, level }) => {
    if (!node) return;
    node.classList.toggle("active", lbv2State.level === level);
    node.classList.toggle("is-left", level < lbv2State.level);
  });

  if (els.lbv2EntriesTopbar && els.lbv2BulkTopbar) {
    els.lbv2EntriesTopbar.classList.toggle("hidden", lbv2State.bulkMode);
    els.lbv2BulkTopbar.classList.toggle("hidden", !lbv2State.bulkMode);
  }

  if (els.lbv2BulkTitle) {
    const count = lbv2State.bulkSelectedEntryIds?.size || 0;
    els.lbv2BulkTitle.textContent = `${count} selected`;
  }
}

function lbv2OpenSheet({ title = "", content = null } = {}) {
  if (!els.lbv2SheetBackdrop || !els.lbv2SheetContent) return;
  if (els.lbv2SheetTitle) els.lbv2SheetTitle.textContent = title;
  els.lbv2SheetContent.innerHTML = "";
  if (content instanceof Node) {
    els.lbv2SheetContent.appendChild(content);
  } else if (typeof content === "string") {
    els.lbv2SheetContent.innerHTML = content;
  }
  els.lbv2SheetBackdrop.classList.remove("hidden");
  els.lbv2SheetBackdrop.setAttribute("aria-hidden", "false");
}

function lbv2CloseSheet() {
  if (!els.lbv2SheetBackdrop || !els.lbv2SheetContent) return;
  els.lbv2SheetBackdrop.classList.add("hidden");
  els.lbv2SheetBackdrop.setAttribute("aria-hidden", "true");
  if (els.lbv2SheetTitle) els.lbv2SheetTitle.textContent = "";
  els.lbv2SheetContent.innerHTML = "";
}

function lbv2OpenDialog({ title = "", body = "", confirmLabel = "Delete", danger = true, onConfirm = null } = {}) {
  if (!els.lbv2DialogBackdrop) return;
  if (els.lbv2DialogTitle) els.lbv2DialogTitle.textContent = title;
  if (els.lbv2DialogBody) els.lbv2DialogBody.textContent = body;
  if (els.lbv2DialogConfirmBtn) {
    els.lbv2DialogConfirmBtn.textContent = confirmLabel;
    els.lbv2DialogConfirmBtn.classList.toggle("lbv2-danger-btn", danger);
    els.lbv2DialogConfirmBtn.classList.toggle("lbv2-secondary-btn", !danger);
  }
  lbv2State.dialog.onConfirm = typeof onConfirm === "function" ? onConfirm : null;
  els.lbv2DialogBackdrop.classList.remove("hidden");
  els.lbv2DialogBackdrop.setAttribute("aria-hidden", "false");
}

function lbv2CloseDialog() {
  if (!els.lbv2DialogBackdrop) return;
  els.lbv2DialogBackdrop.classList.add("hidden");
  els.lbv2DialogBackdrop.setAttribute("aria-hidden", "true");
  lbv2State.dialog.onConfirm = null;
}

function lbv2CreateSheetButton({ label, sublabel = "", className = "", onClick }) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `lbv2-sheet-item ${className}`.trim();

  const left = document.createElement("span");
  left.textContent = label;
  button.appendChild(left);

  if (sublabel) {
    const small = document.createElement("small");
    small.className = "lbv2-sheet-sub";
    small.textContent = sublabel;
    button.appendChild(small);
  }

  button.addEventListener("click", () => {
    if (typeof onClick === "function") onClick();
  });

  return button;
}

function lbv2ShowSaveIndicator(text = "Saved") {
  if (!els.lbv2SaveIndicator) return;
  if (lbv2State.saveIndicatorTimer) clearTimeout(lbv2State.saveIndicatorTimer);
  els.lbv2SaveIndicator.textContent = text;
  els.lbv2SaveIndicator.classList.add("show");
  lbv2State.saveIndicatorTimer = setTimeout(() => {
    els.lbv2SaveIndicator?.classList.remove("show");
  }, 1500);
}

function lbv2UpdateSwitchAria() {
  if (!els.lbv2Panel) return;
  const switchInputs = els.lbv2Panel.querySelectorAll('input[role="switch"]');
  switchInputs.forEach((input) => {
    input.setAttribute("aria-checked", input.checked ? "true" : "false");
  });
}

function lbv2ApplyEntrySort(entries = []) {
  const list = Array.isArray(entries) ? entries.slice() : [];
  if (lbv2State.entrySortBy === LOREBOOK_V2_SORT_OPTIONS.ALPHABETICAL) {
    list.sort((a, b) => String(a.title || "").localeCompare(String(b.title || "")));
  } else if (lbv2State.entrySortBy === LOREBOOK_V2_SORT_OPTIONS.PRIORITY) {
    list.sort((a, b) => Number(b.priority || 0) - Number(a.priority || 0));
  } else if (lbv2State.entrySortBy === LOREBOOK_V2_SORT_OPTIONS.TOKEN_COUNT) {
    list.sort((a, b) => estimateTokenCount(b.content || "") - estimateTokenCount(a.content || ""));
  } else if (lbv2State.entrySortBy === LOREBOOK_V2_SORT_OPTIONS.RECENTLY_MODIFIED) {
    list.sort((a, b) => new Date(b.updatedAt || 0) - new Date(a.updatedAt || 0));
  } else {
    list.sort((a, b) => Number(a.insertionOrder || 0) - Number(b.insertionOrder || 0));
  }
  return list;
}

function lbv2ApplyLorebookSort(lorebooks = []) {
  const list = Array.isArray(lorebooks) ? lorebooks.slice() : [];
  if (lbv2State.lorebookSortBy === LOREBOOK_V2_SORT_OPTIONS.ALPHABETICAL) {
    list.sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")));
  } else if (lbv2State.lorebookSortBy === LOREBOOK_V2_SORT_OPTIONS.DATE_CREATED) {
    list.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
  } else if (lbv2State.lorebookSortBy === LOREBOOK_V2_SORT_OPTIONS.ENTRY_COUNT) {
    list.sort((a, b) => (b.entries?.length || 0) - (a.entries?.length || 0));
  } else {
    list.sort((a, b) => new Date(b.updatedAt || 0) - new Date(a.updatedAt || 0));
  }
  return list;
}

function lbv2CreateEmptyEntry(lorebook) {
  const defaultsEntry = lorebook?.defaultEntrySettings || {};
  const nowIso = new Date().toISOString();
  return normalizeLorebookEntry({
    id: crypto.randomUUID(),
    title: "Untitled Entry",
    keys: [],
    secondary_keys: [],
    activation_logic: "ANY",
    content: "",
    insertion_order: Number.isFinite(Number(defaultsEntry.insertionOrder)) ? Number(defaultsEntry.insertionOrder) : 100,
    insertion_position: defaultsEntry.insertionPosition || "after_system",
    scan_depth: Number.isFinite(Number(defaultsEntry.scanDepth)) ? Number(defaultsEntry.scanDepth) : 10,
    token_budget: 0,
    priority: Number.isFinite(Number(defaultsEntry.insertionOrder)) ? Number(defaultsEntry.insertionOrder) : 100,
    case_sensitive: Boolean(defaultsEntry.caseSensitive),
    enabled: defaultsEntry.enabled !== false,
    strategy: "normal",
    created_at: nowIso,
    updated_at: nowIso,
  });
}

function lbv2PersistLorebook(lorebook, { savedText = "Saved" } = {}) {
  if (!lorebook) return;
  lorebook.updatedAt = new Date().toISOString();
  const index = state.lorebooks.findIndex((item) => item.id === lorebook.id);
  if (index >= 0) {
    state.lorebooks[index] = normalizeLorebookRecord(lorebook);
  } else {
    state.lorebooks.unshift(normalizeLorebookRecord(lorebook));
  }

  const normalizedBook = state.lorebooks.find((item) => item.id === lorebook.id);
  if (normalizedBook) {
    const chats = Array.isArray(state.chats) ? state.chats : [];
    chats.forEach((chat) => {
      if (!chat?.id) return;
      if (!Array.isArray(chat.lorebookIds)) chat.lorebookIds = [];
      const set = new Set(chat.lorebookIds.map((id) => String(id || "")).filter(Boolean));

      const attached = lbv2IsLorebookAttachedToChat(normalizedBook, chat);
      if (attached) set.add(normalizedBook.id);
      else set.delete(normalizedBook.id);

      chat.lorebookIds = [...set];
    });
  }

  persistState();
  renderDrawerLorebooks();
  lbv2ShowSaveIndicator(savedText);
}

function lbv2ScheduleAutosave() {
  if (lbv2State.suppressEditorEvents) return;
  if (lbv2State.level !== 3) return;
  if (lbv2State.autosaveTimer) clearTimeout(lbv2State.autosaveTimer);
  lbv2State.autosaveTimer = setTimeout(() => {
    lbv2SaveEntryFromEditor({ silent: true });
  }, 500);
}

function initLorebookV2() {
  if (!LOREBOOK_V2_ENABLED || !els.lbv2Panel || !els.lbv2LorebookList) return;

  lbv2State.level = 1;
  lbv2State.activeLorebookId = "";
  lbv2State.activeEntryId = "";
  lbv2State.bulkMode = false;
  lbv2State.bulkSelectedEntryIds = new Set();
  lbv2SyncStack();
  lbv2UpdateSwitchAria();
  renderLorebooksV2();
}

function openLibraryLorebooksPanel() {
  setActiveLibraryPanel("libraryLorebooksPanel");
}

function bindLorebookV2Events() {
  if (!LOREBOOK_V2_ENABLED || !els.lbv2Panel || !els.lbv2Stack) return;

  const goBack = () => lbv2NavigateBack();
  els.lbv2BackLibraryBtn?.addEventListener("click", goBack);
  els.lbv2EntriesBackBtn?.addEventListener("click", goBack);
  els.lbv2EditorBackBtn?.addEventListener("click", goBack);

  const openCreate = () => lbv2OpenCreateLorebookSheet();
  els.lbv2CreateLorebookTopBtn?.addEventListener("click", openCreate);
  els.lbv2EmptyCreateBtn?.addEventListener("click", openCreate);
  els.lbv2LorebookFab?.addEventListener("click", openCreate);

  els.lbv2LorebookSearchInput?.addEventListener("input", () => {
    lbv2State.lorebookSearchQuery = String(els.lbv2LorebookSearchInput.value || "");
    renderLorebooksV2();
  });

  els.lbv2EntrySearchInput?.addEventListener("input", () => {
    lbv2State.entrySearchQuery = String(els.lbv2EntrySearchInput.value || "");
    lbv2RenderEntryList();
  });

  els.lbv2LorebookOverflowBtn?.addEventListener("click", lbv2OpenLorebookOverflowSheet);
  els.lbv2EntriesOverflowBtn?.addEventListener("click", lbv2OpenEntriesOverflowSheet);
  els.lbv2EditorOverflowBtn?.addEventListener("click", lbv2OpenEditorOverflowSheet);

  els.lbv2LorebookSettingsBtn?.addEventListener("click", lbv2OpenSettingsModal);
  els.lbv2SettingsBackBtn?.addEventListener("click", lbv2CloseSettingsModal);

  els.lbv2EntryNewBtn?.addEventListener("click", lbv2CreateEntryAndOpenEditor);
  els.lbv2EmptyEntryCreateBtn?.addEventListener("click", lbv2CreateEntryAndOpenEditor);

  els.lbv2BulkCancelBtn?.addEventListener("click", () => {
    lbv2State.bulkMode = false;
    lbv2State.bulkSelectedEntryIds = new Set();
    lbv2SyncStack();
    lbv2RenderEntryList();
  });

  els.lbv2BulkActionsBtn?.addEventListener("click", lbv2OpenBulkActionsSheet);

  els.lbv2SheetBackdrop?.addEventListener("click", (event) => {
    if (event.target === els.lbv2SheetBackdrop) lbv2CloseSheet();
  });

  els.lbv2DialogBackdrop?.addEventListener("click", (event) => {
    if (event.target === els.lbv2DialogBackdrop) lbv2CloseDialog();
  });

  els.lbv2DialogCancelBtn?.addEventListener("click", lbv2CloseDialog);
  els.lbv2DialogConfirmBtn?.addEventListener("click", () => {
    const fn = lbv2State.dialog.onConfirm;
    lbv2CloseDialog();
    if (typeof fn === "function") fn();
  });

  els.lbv2EntryImportInput?.addEventListener("change", lbv2ImportEntriesIntoActiveLorebook);

  els.lbv2EntryContentInput?.addEventListener("input", () => {
    lbv2UpdateEditorDerivedUi();
    lbv2ScheduleAutosave();
  });
  els.lbv2EntryNameInput?.addEventListener("input", lbv2ScheduleAutosave);
  els.lbv2EntryCommentInput?.addEventListener("input", lbv2ScheduleAutosave);
  els.lbv2PrimaryKeysInput?.addEventListener("input", () => {
    lbv2UpdateEditorDerivedUi();
    lbv2ScheduleAutosave();
  });
  els.lbv2SecondaryKeysInput?.addEventListener("input", lbv2ScheduleAutosave);

  [
    els.lbv2ActivationAnyInput,
    els.lbv2ActivationAllInput,
    els.lbv2ActivationAndOrInput,
    els.lbv2ActivationNotInput,
    els.lbv2CaseSensitiveInput,
    els.lbv2WholeWordsInput,
    els.lbv2RegexInput,
    els.lbv2ScanDepthInput,
    els.lbv2PerEntryTokenBudgetInput,
    els.lbv2InsertionPositionInput,
    els.lbv2InsertionDepthInput,
    els.lbv2PriorityInput,
    els.lbv2EnabledInput,
    els.lbv2PreventRecursionInput,
    els.lbv2ExcludeFromRecursionInput,
    els.lbv2ActivationProbabilitySlider,
    els.lbv2ActivationProbabilityInput,
    els.lbv2GroupInput,
    els.lbv2GroupWeightInput,
    els.lbv2CategoryInput,
    els.lbv2AutomationIdInput,
    els.lbv2StickyInput,
    els.lbv2CooldownInput,
    els.lbv2DelayInput,
    els.lbv2TurnStartInput,
    els.lbv2TurnEndInput,
  ].forEach((node) => {
    node?.addEventListener("change", () => {
      lbv2UpdateEditorDerivedUi();
      lbv2ScheduleAutosave();
      lbv2UpdateSwitchAria();
    });
    node?.addEventListener("input", () => {
      lbv2UpdateEditorDerivedUi();
      lbv2ScheduleAutosave();
    });
  });

  const roleRadios = els.lbv2Panel.querySelectorAll('input[name="lbv2MessageRole"]');
  roleRadios.forEach((radio) => radio.addEventListener("change", lbv2ScheduleAutosave));

  const strategyRadios = els.lbv2Panel.querySelectorAll('input[name="lbv2Strategy"]');
  strategyRadios.forEach((radio) => radio.addEventListener("change", () => {
    lbv2UpdateEditorDerivedUi();
    lbv2ScheduleAutosave();
  }));

  els.lbv2InsertionPositionBtn?.addEventListener("click", () => lbv2OpenPickerSheet("insertion-position"));
  els.lbv2CategoryPickerBtn?.addEventListener("click", () => lbv2OpenPickerSheet("category"));
  els.lbv2SettingsDefaultPositionBtn?.addEventListener("click", () => lbv2OpenPickerSheet("default-position"));
  els.lbv2LinkEntryBtn?.addEventListener("click", lbv2OpenLinkEntryPicker);

  els.lbv2TagInput?.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === ",") {
      event.preventDefault();
      lbv2AddTagFromInput();
    }
  });
  els.lbv2TagInput?.addEventListener("blur", lbv2AddTagFromInput);
  els.lbv2TagInput?.addEventListener("input", lbv2RenderTagSuggestions);

  els.lbv2ActivationProbabilitySlider?.addEventListener("input", () => {
    if (!els.lbv2ActivationProbabilityInput) return;
    els.lbv2ActivationProbabilityInput.value = els.lbv2ActivationProbabilitySlider.value;
  });
  els.lbv2ActivationProbabilityInput?.addEventListener("input", () => {
    if (!els.lbv2ActivationProbabilitySlider) return;
    const value = Math.min(100, Math.max(0, Number.parseInt(els.lbv2ActivationProbabilityInput.value || "100", 10) || 100));
    els.lbv2ActivationProbabilityInput.value = String(value);
    els.lbv2ActivationProbabilitySlider.value = String(value);
  });

  els.lbv2RegexInput?.addEventListener("change", () => {
    if (els.lbv2RegexHelper) {
      els.lbv2RegexHelper.classList.toggle("hidden", !els.lbv2RegexInput.checked);
    }
  });

  els.lbv2SettingsRecursiveInput?.addEventListener("change", () => {
    els.lbv2SettingsRecursionDepthField?.classList.toggle("hidden", !els.lbv2SettingsRecursiveInput.checked);
    lbv2SaveSettingsModal();
  });

  els.lbv2SettingsNameInput?.addEventListener("input", lbv2SaveSettingsModal);
  els.lbv2SettingsDescriptionInput?.addEventListener("input", lbv2SaveSettingsModal);
  els.lbv2SettingsTokenBudgetInput?.addEventListener("change", lbv2SaveSettingsModal);
  els.lbv2SettingsDefaultScanDepthInput?.addEventListener("change", lbv2SaveSettingsModal);
  els.lbv2SettingsDefaultOrderInput?.addEventListener("change", lbv2SaveSettingsModal);
  els.lbv2SettingsDefaultEnabledInput?.addEventListener("change", lbv2SaveSettingsModal);
  els.lbv2SettingsDefaultCaseInput?.addEventListener("change", lbv2SaveSettingsModal);
  els.lbv2SettingsRecursionDepthInput?.addEventListener("change", lbv2SaveSettingsModal);

  els.lbv2AttachmentCharacterSelect?.addEventListener("change", lbv2SaveSettingsModal);
  const attachmentScopeRadios = els.lbv2Panel.querySelectorAll('input[name="lbv2AttachmentScope"]');
  attachmentScopeRadios.forEach((radio) => {
    radio.addEventListener("change", () => {
      lbv2UpdateAttachmentVisibility();
      lbv2SaveSettingsModal();
    });
  });

  els.lbv2AttachChatBtn?.addEventListener("click", lbv2OpenAttachChatPicker);

  els.lbv2ExportJsonBtn?.addEventListener("click", () => {
    const lorebook = lbv2GetActiveLorebook();
    if (!lorebook) return;
    exportLorebook(lorebook.id);
    lbv2CloseSheet();
  });
  els.lbv2ExportStBtn?.addEventListener("click", lbv2ExportActiveLorebookAsSillyTavern);
  els.lbv2ImportEntriesBtn?.addEventListener("click", () => els.lbv2EntryImportInput?.click());

  els.lbv2DisableAllEntriesBtn?.addEventListener("click", () => {
    const lorebook = lbv2GetActiveLorebook();
    if (!lorebook) return;
    lorebook.entries = (lorebook.entries || []).map((entry) => ({ ...entry, enabled: false, updatedAt: new Date().toISOString() }));
    lbv2PersistLorebook(lorebook);
    lbv2RenderEntryList();
  });

  els.lbv2EnableAllEntriesBtn?.addEventListener("click", () => {
    const lorebook = lbv2GetActiveLorebook();
    if (!lorebook) return;
    lorebook.entries = (lorebook.entries || []).map((entry) => ({ ...entry, enabled: true, updatedAt: new Date().toISOString() }));
    lbv2PersistLorebook(lorebook);
    lbv2RenderEntryList();
  });

  els.lbv2DeleteLorebookBtn?.addEventListener("click", () => {
    const lorebook = lbv2GetActiveLorebook();
    if (!lorebook) return;
    const count = Array.isArray(lorebook.entries) ? lorebook.entries.length : 0;
    lbv2OpenDialog({
      title: `Delete ${lorebook.name}?`,
      body: `This will remove all ${count} entries. This cannot be undone.`,
      confirmLabel: "Delete",
      danger: true,
      onConfirm: () => {
        removeLorebook(lorebook.id);
        lbv2CloseSettingsModal();
        lbv2NavigateTo(1);
      },
    });
  });

  els.lbv2Stack.addEventListener("pointerdown", (event) => {
    if (event.pointerType === "mouse" && event.button !== 0) return;
    if (event.clientX > 28) return;
    lbv2State.gesture.pointerId = event.pointerId;
    lbv2State.gesture.startX = event.clientX;
    lbv2State.gesture.startY = event.clientY;
    lbv2State.gesture.tracking = true;
  });

  els.lbv2Stack.addEventListener("pointermove", (event) => {
    if (!lbv2State.gesture.tracking || event.pointerId !== lbv2State.gesture.pointerId) return;
    const dx = event.clientX - lbv2State.gesture.startX;
    const dy = Math.abs(event.clientY - lbv2State.gesture.startY);
    if (dx > 72 && dy < 36) {
      lbv2State.gesture.tracking = false;
      lbv2NavigateBack();
    }
  });

  const stopGesture = (event) => {
    if (event.pointerId !== lbv2State.gesture.pointerId) return;
    lbv2State.gesture.pointerId = null;
    lbv2State.gesture.tracking = false;
  };

  els.lbv2Stack.addEventListener("pointerup", stopGesture);
  els.lbv2Stack.addEventListener("pointercancel", stopGesture);

  els.lbv2Panel.querySelectorAll(".lbv2-accordion-header").forEach((header) => {
    header.addEventListener("click", () => {
      const accordion = header.closest(".lbv2-accordion");
      if (!accordion) return;
      accordion.classList.toggle("open");
      const expanded = accordion.classList.contains("open");
      header.setAttribute("aria-expanded", expanded ? "true" : "false");
    });
  });

  els.lbv2Panel.querySelectorAll(".lbv2-stepper-btn").forEach((button) => {
    const applyStep = () => {
      const targetId = button.dataset.stepTarget;
      const step = Number.parseInt(button.dataset.step || "1", 10) || 1;
      if (!targetId) return;
      const input = document.getElementById(targetId);
      if (!input) return;
      const min = Number.isFinite(Number(input.min)) ? Number(input.min) : null;
      const max = Number.isFinite(Number(input.max)) ? Number(input.max) : null;
      let value = Number.parseInt(input.value || "0", 10) || 0;
      value += step;
      if (min !== null) value = Math.max(min, value);
      if (max !== null) value = Math.min(max, value);
      input.value = String(value);
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
    };

    button.addEventListener("click", applyStep);

    button.addEventListener("pointerdown", (event) => {
      if (event.pointerType === "mouse" && event.button !== 0) return;
      if (lbv2State.stepperHold.timer) clearTimeout(lbv2State.stepperHold.timer);
      if (lbv2State.stepperHold.interval) clearInterval(lbv2State.stepperHold.interval);
      lbv2State.stepperHold.timer = setTimeout(() => {
        lbv2State.stepperHold.interval = setInterval(applyStep, 120);
      }, 360);
    });
  });

  const clearStepperHold = () => {
    if (lbv2State.stepperHold.timer) clearTimeout(lbv2State.stepperHold.timer);
    if (lbv2State.stepperHold.interval) clearInterval(lbv2State.stepperHold.interval);
    lbv2State.stepperHold.timer = null;
    lbv2State.stepperHold.interval = null;
  };

  document.addEventListener("pointerup", clearStepperHold);
  document.addEventListener("pointercancel", clearStepperHold);
}

function renderLorebooksV2() {
  if (!LOREBOOK_V2_ENABLED || !els.lbv2LorebookList) return;

  state.lorebooks = (Array.isArray(state.lorebooks) ? state.lorebooks : [])
    .map((book) => withLorebookTimestamps(normalizeLorebookRecord(book), book));

  lbv2RenderLorebookList();
  lbv2RenderEntryList();
  lbv2RenderEditor();
  lbv2SyncStack();
  lbv2UpdateSwitchAria();
}

function lbv2GetLorebookChatAttachmentCount(book) {
  if (!book) return 0;
  return (Array.isArray(state.chats) ? state.chats : []).filter((chat) => lbv2IsLorebookAttachedToChat(book, chat)).length;
}

function lbv2RenderLorebookList() {
  if (!els.lbv2LorebookList) return;

  const query = String(lbv2State.lorebookSearchQuery || "").trim().toLowerCase();
  const currentChat = getCurrentChat();

  let lorebooks = lbv2ApplyLorebookSort(
    (Array.isArray(state.lorebooks) ? state.lorebooks : [])
      .map((book) => withLorebookTimestamps(book, book))
      .filter((book) => {
        if (!query) return true;
        return `${book.name || ""} ${book.description || ""}`.toLowerCase().includes(query);
      })
  );

  els.lbv2LorebookList.innerHTML = "";

  if (!lorebooks.length) {
    els.lbv2LorebookEmptyState?.classList.remove("hidden");
    return;
  }

  els.lbv2LorebookEmptyState?.classList.add("hidden");

  lorebooks.forEach((book) => {
    const attached = lbv2IsLorebookAttachedToChat(book, currentChat);
    const tokenTotal = countLorebookTokens(book);
    const attachedCount = lbv2GetLorebookChatAttachmentCount(book);

    const card = document.createElement("article");
    card.className = "lbv2-lorebook-card";
    card.dataset.lorebookId = book.id;
    card.innerHTML = `
      <div class="lbv2-lorebook-row1">
        <span class="lbv2-color-dot" style="background:${escapeAttribute(book.color || LOREBOOK_V2_COLORS[0])};"></span>
        <h4 class="lbv2-lorebook-name">${escapeHtml(book.name || "Untitled lorebook")}</h4>
        <button type="button" class="lbv2-toggle${attached ? " is-on" : ""}" aria-label="Toggle lorebook in current chat"></button>
      </div>
      <p class="lbv2-lorebook-desc">${escapeHtml(book.description || "No description")}</p>
      <div class="lbv2-lorebook-meta">${book.entries?.length || 0} entries • ${tokenTotal.toLocaleString()} tk • ${attachedCount} chat(s)</div>
    `;

    const toggle = card.querySelector(".lbv2-toggle");
    toggle?.addEventListener("click", (event) => {
      event.stopPropagation();
      lbv2ToggleLorebookAttachment(book.id);
    });

    card.addEventListener("click", () => {
      lbv2State.activeLorebookId = book.id;
      lbv2State.activeEntryId = "";
      lbv2State.entrySearchQuery = "";
      if (els.lbv2EntrySearchInput) els.lbv2EntrySearchInput.value = "";
      lbv2State.bulkMode = false;
      lbv2State.bulkSelectedEntryIds = new Set();
      lbv2NavigateTo(2);
      lbv2RenderEntryList();
    });

    const openContext = () => lbv2OpenLorebookContextSheet(book);
    card.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      openContext();
    });

    let pressTimer = null;
    let startX = 0;
    let startY = 0;
    const clearPress = () => {
      if (pressTimer) clearTimeout(pressTimer);
      pressTimer = null;
    };

    card.addEventListener("pointerdown", (event) => {
      if (event.pointerType === "mouse" && event.button !== 0) return;
      startX = event.clientX;
      startY = event.clientY;
      clearPress();
      pressTimer = setTimeout(() => {
        openContext();
      }, 520);
    });

    card.addEventListener("pointermove", (event) => {
      if (!pressTimer) return;
      if (Math.abs(event.clientX - startX) > 10 || Math.abs(event.clientY - startY) > 10) {
        clearPress();
      }
    });

    card.addEventListener("pointerup", clearPress);
    card.addEventListener("pointercancel", clearPress);

    els.lbv2LorebookList.appendChild(card);
  });
}

function lbv2OpenLorebookContextSheet(book) {
  if (!book) return;

  const container = document.createElement("div");
  container.className = "lbv2-sheet-content";

  container.appendChild(lbv2CreateSheetButton({
    label: "Rename",
    onClick: () => {
      lbv2OpenSheet({
        title: "Rename lorebook",
        content: (() => {
          const form = document.createElement("div");
          form.className = "lbv2-sheet-form";
          const input = document.createElement("input");
          input.type = "text";
          input.value = book.name || "";
          input.placeholder = "Lorebook name";
          const saveBtn = document.createElement("button");
          saveBtn.type = "button";
          saveBtn.className = "lbv2-primary-btn";
          saveBtn.textContent = "Save";
          saveBtn.addEventListener("click", () => {
            const value = input.value.trim() || "Untitled lorebook";
            book.name = value;
            lbv2PersistLorebook(book);
            renderLorebooksV2();
            lbv2CloseSheet();
          });
          form.append(input, saveBtn);
          setTimeout(() => input.focus(), 40);
          return form;
        })(),
      });
    },
  }));

  container.appendChild(lbv2CreateSheetButton({
    label: "Duplicate",
    onClick: () => {
      const nowIso = new Date().toISOString();
      const duplicated = normalizeLorebookRecord({
        ...book,
        id: crypto.randomUUID(),
        name: `${book.name || "Lorebook"} (Copy)`,
        createdAt: nowIso,
        updatedAt: nowIso,
        entries: (Array.isArray(book.entries) ? book.entries : []).map((entry, index) => ({
          ...entry,
          id: crypto.randomUUID(),
          insertion_order: index,
          created_at: nowIso,
          updated_at: nowIso,
        })),
      });
      state.lorebooks.unshift(duplicated);
      persistState();
      renderLorebooksV2();
      renderDrawerLorebooks();
      lbv2CloseSheet();
      showToast("Lorebook duplicated", "success");
    },
  }));

  container.appendChild(lbv2CreateSheetButton({
    label: "Export (JSON)",
    onClick: () => {
      exportLorebook(book.id);
      lbv2CloseSheet();
    },
  }));

  container.appendChild(lbv2CreateSheetButton({
    label: "Delete",
    className: "danger",
    onClick: () => {
      lbv2CloseSheet();
      lbv2OpenDialog({
        title: `Delete ${book.name}?`,
        body: `This will remove all ${book.entries?.length || 0} entries. This cannot be undone.`,
        confirmLabel: "Delete",
        danger: true,
        onConfirm: () => {
          removeLorebook(book.id);
          lbv2NavigateTo(1);
        },
      });
    },
  }));

  lbv2OpenSheet({
    title: book.name || "Lorebook",
    content: container,
  });
}

function lbv2OpenCreateLorebookSheet() {
  const form = document.createElement("div");
  form.className = "lbv2-sheet-form";

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.placeholder = "Lorebook name";

  const descInput = document.createElement("textarea");
  descInput.rows = 3;
  descInput.placeholder = "Description (optional)";

  const colorWrap = document.createElement("div");
  colorWrap.className = "lbv2-color-dot-row";
  let selectedColor = LOREBOOK_V2_COLORS[0];
  LOREBOOK_V2_COLORS.forEach((color, index) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `lbv2-color-choice${index === 0 ? " active" : ""}`;
    btn.style.background = color;
    btn.setAttribute("aria-label", `Color ${index + 1}`);
    btn.addEventListener("click", () => {
      selectedColor = color;
      colorWrap.querySelectorAll(".lbv2-color-choice").forEach((node) => node.classList.remove("active"));
      btn.classList.add("active");
    });
    colorWrap.appendChild(btn);
  });

  const createBtn = document.createElement("button");
  createBtn.type = "button";
  createBtn.className = "lbv2-primary-btn";
  createBtn.textContent = "Create";

  createBtn.addEventListener("click", () => {
    const nowIso = new Date().toISOString();
    const lorebook = normalizeLorebookRecord({
      id: crypto.randomUUID(),
      name: nameInput.value.trim() || "Untitled lorebook",
      description: descInput.value.trim(),
      color: selectedColor,
      createdAt: nowIso,
      updatedAt: nowIso,
      entries: [],
      defaultEntrySettings: {
        insertionPosition: "after_system",
        scanDepth: 10,
        insertionOrder: 100,
        enabled: true,
        caseSensitive: false,
      },
      attachmentScope: "specific",
      attachedChatIds: getCurrentChat()?.id ? [getCurrentChat().id] : [],
      recursiveScanning: true,
      maxRecursionDepth: 3,
    });

    state.lorebooks.unshift(lorebook);

    const chat = getCurrentChat();
    if (chat?.id) {
      chat.lorebookIds = [...new Set([...(chat.lorebookIds || []), lorebook.id])];
      syncCurrentChatFromState();
    }

    persistState();
    renderDrawerLorebooks();
    lbv2CloseSheet();

    lbv2State.activeLorebookId = lorebook.id;
    lbv2NavigateTo(2);
    renderLorebooksV2();
  });

  form.append(nameInput, descInput, colorWrap, createBtn);

  lbv2OpenSheet({
    title: "New Lorebook",
    content: form,
  });

  setTimeout(() => nameInput.focus(), 40);
}

function lbv2OpenLorebookOverflowSheet() {
  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  content.appendChild(lbv2CreateSheetButton({
    label: "Import Lorebook",
    onClick: () => {
      lbv2CloseSheet();
      els.lorebookImportInput?.click();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Export All Lorebooks",
    onClick: () => {
      exportLorebooks();
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Sort By",
    sublabel: "Alphabetical, Date Modified, Date Created, Entry Count",
    onClick: () => {
      lbv2OpenLorebookSortPicker();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Bulk Delete Mode",
    className: "warning",
    onClick: () => {
      lbv2OpenLorebookBulkDeleteSheet();
    },
  }));

  lbv2OpenSheet({ title: "Lorebook options", content });
}

function lbv2OpenLorebookSortPicker() {
  const options = [
    { key: LOREBOOK_V2_SORT_OPTIONS.ALPHABETICAL, label: "Alphabetical" },
    { key: LOREBOOK_V2_SORT_OPTIONS.DATE_MODIFIED, label: "Date Modified" },
    { key: LOREBOOK_V2_SORT_OPTIONS.DATE_CREATED, label: "Date Created" },
    { key: LOREBOOK_V2_SORT_OPTIONS.ENTRY_COUNT, label: "Entry Count" },
  ];

  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  options.forEach((option) => {
    content.appendChild(lbv2CreateSheetButton({
      label: option.label,
      sublabel: lbv2State.lorebookSortBy === option.key ? "Selected" : "",
      onClick: () => {
        lbv2State.lorebookSortBy = option.key;
        renderLorebooksV2();
        lbv2CloseSheet();
      },
    }));
  });

  lbv2OpenSheet({ title: "Sort Lorebooks", content });
}

function lbv2OpenLorebookBulkDeleteSheet() {
  const books = Array.isArray(state.lorebooks) ? state.lorebooks : [];
  if (!books.length) {
    showToast("No lorebooks to delete", "error");
    lbv2CloseSheet();
    return;
  }

  const form = document.createElement("div");
  form.className = "lbv2-sheet-form";
  const selection = new Set();

  books.forEach((book) => {
    const row = document.createElement("label");
    row.className = "lbv2-radio-row";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) selection.add(book.id);
      else selection.delete(book.id);
      deleteBtn.textContent = `Delete Selected (${selection.size})`;
    });
    const span = document.createElement("span");
    span.textContent = `${book.name || "Untitled"} (${book.entries?.length || 0})`;
    row.append(checkbox, span);
    form.appendChild(row);
  });

  const deleteBtn = document.createElement("button");
  deleteBtn.type = "button";
  deleteBtn.className = "lbv2-danger-btn";
  deleteBtn.textContent = "Delete Selected (0)";
  deleteBtn.addEventListener("click", () => {
    if (!selection.size) return;
    lbv2OpenDialog({
      title: `Delete ${selection.size} lorebook${selection.size === 1 ? "" : "s"}?`,
      body: "This action cannot be undone.",
      confirmLabel: "Delete",
      danger: true,
      onConfirm: () => {
        [...selection].forEach((id) => removeLorebook(id));
        lbv2CloseSheet();
        renderLorebooksV2();
      },
    });
  });

  form.appendChild(deleteBtn);
  lbv2OpenSheet({ title: "Bulk Delete", content: form });
}

function lbv2OpenEntriesOverflowSheet() {
  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  content.appendChild(lbv2CreateSheetButton({
    label: "Sort By",
    onClick: lbv2OpenEntrySortPicker,
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Filter",
    onClick: lbv2OpenEntryFilterPicker,
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: lbv2State.bulkMode ? "Exit Bulk Select Mode" : "Bulk Select Mode",
    onClick: () => {
      lbv2State.bulkMode = !lbv2State.bulkMode;
      if (!lbv2State.bulkMode) {
        lbv2State.bulkSelectedEntryIds = new Set();
      }
      lbv2SyncStack();
      lbv2RenderEntryList();
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Import Entries",
    onClick: () => {
      lbv2CloseSheet();
      els.lbv2EntryImportInput?.click();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Export Entries",
    onClick: () => {
      const lorebook = lbv2GetActiveLorebook();
      if (!lorebook) return;
      downloadJson(`${slugify(lorebook.name || "lorebook")}-entries.json`, lorebook.entries || []);
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Collapse/Expand All Previews",
    onClick: () => {
      lbv2State.expandedEntrySwipeId = lbv2State.expandedEntrySwipeId ? "" : "__expand__";
      lbv2RenderEntryList();
      lbv2CloseSheet();
    },
  }));

  lbv2OpenSheet({ title: "Entry options", content });
}

function lbv2OpenEntrySortPicker() {
  const options = [
    { key: LOREBOOK_V2_SORT_OPTIONS.MANUAL, label: "Manual (drag reorder)" },
    { key: LOREBOOK_V2_SORT_OPTIONS.ALPHABETICAL, label: "Alphabetical" },
    { key: LOREBOOK_V2_SORT_OPTIONS.PRIORITY, label: "Insertion Order/Priority" },
    { key: LOREBOOK_V2_SORT_OPTIONS.TOKEN_COUNT, label: "Token Count" },
    { key: LOREBOOK_V2_SORT_OPTIONS.RECENTLY_MODIFIED, label: "Recently Modified" },
  ];
  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";
  options.forEach((option) => {
    content.appendChild(lbv2CreateSheetButton({
      label: option.label,
      sublabel: lbv2State.entrySortBy === option.key ? "Selected" : "",
      onClick: () => {
        lbv2State.entrySortBy = option.key;
        lbv2RenderEntryList();
        lbv2CloseSheet();
      },
    }));
  });
  lbv2OpenSheet({ title: "Sort entries", content });
}

function lbv2OpenEntryFilterPicker() {
  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook) return;

  const tags = [...new Set((lorebook.entries || []).flatMap((entry) => Array.isArray(entry.tags) ? entry.tags : []))]
    .map((tag) => String(tag || "").trim())
    .filter(Boolean);

  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  [
    { value: "all", label: "All" },
    { value: "enabled", label: "Enabled Only" },
    { value: "disabled", label: "Disabled Only" },
    { value: "constant", label: "Constant Only" },
  ].forEach((option) => {
    content.appendChild(lbv2CreateSheetButton({
      label: option.label,
      sublabel: lbv2State.entryFilter === option.value ? "Selected" : "",
      onClick: () => {
        lbv2State.entryFilter = option.value;
        lbv2State.entryTagFilter = "";
        lbv2RenderEntryList();
        lbv2CloseSheet();
      },
    }));
  });

  if (tags.length) {
    tags.forEach((tag) => {
      content.appendChild(lbv2CreateSheetButton({
        label: `Tag: ${tag}`,
        sublabel: lbv2State.entryTagFilter === tag ? "Selected" : "",
        onClick: () => {
          lbv2State.entryFilter = "tag";
          lbv2State.entryTagFilter = tag;
          lbv2RenderEntryList();
          lbv2CloseSheet();
        },
      }));
    });
  }

  lbv2OpenSheet({ title: "Filter entries", content });
}

function lbv2OpenBulkActionsSheet() {
  const selected = [...(lbv2State.bulkSelectedEntryIds || new Set())];
  if (!selected.length) {
    showToast("No entries selected", "error");
    return;
  }

  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  content.appendChild(lbv2CreateSheetButton({
    label: "Enable Selected",
    onClick: () => {
      const lorebook = lbv2GetActiveLorebook();
      if (!lorebook) return;
      lorebook.entries = (lorebook.entries || []).map((entry) => selected.includes(entry.id) ? { ...entry, enabled: true, updatedAt: new Date().toISOString() } : entry);
      lbv2PersistLorebook(lorebook);
      lbv2RenderEntryList();
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Disable Selected",
    onClick: () => {
      const lorebook = lbv2GetActiveLorebook();
      if (!lorebook) return;
      lorebook.entries = (lorebook.entries || []).map((entry) => selected.includes(entry.id) ? { ...entry, enabled: false, updatedAt: new Date().toISOString() } : entry);
      lbv2PersistLorebook(lorebook);
      lbv2RenderEntryList();
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Set Category",
    onClick: () => {
      const wrap = document.createElement("div");
      wrap.className = "lbv2-sheet-form";
      const input = document.createElement("input");
      input.type = "text";
      input.placeholder = "Category";
      const saveBtn = document.createElement("button");
      saveBtn.type = "button";
      saveBtn.className = "lbv2-primary-btn";
      saveBtn.textContent = "Apply";
      saveBtn.addEventListener("click", () => {
        const lorebook = lbv2GetActiveLorebook();
        if (!lorebook) return;
        const value = input.value.trim();
        lorebook.entries = (lorebook.entries || []).map((entry) => selected.includes(entry.id)
          ? { ...entry, category: value, updatedAt: new Date().toISOString() }
          : entry);
        lbv2PersistLorebook(lorebook);
        lbv2RenderEntryList();
        lbv2CloseSheet();
      });
      wrap.append(input, saveBtn);
      lbv2OpenSheet({ title: "Set category", content: wrap });
      setTimeout(() => input.focus(), 40);
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Move to Another Lorebook",
    onClick: () => lbv2MoveEntriesToLorebook(selected),
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Delete Selected",
    className: "danger",
    onClick: () => {
      lbv2OpenDialog({
        title: `Delete ${selected.length} entries?`,
        body: "This cannot be undone.",
        confirmLabel: "Delete",
        danger: true,
        onConfirm: () => {
          const lorebook = lbv2GetActiveLorebook();
          if (!lorebook) return;
          lorebook.entries = (lorebook.entries || []).filter((entry) => !selected.includes(entry.id));
          lorebook.entries = lorebook.entries.map((entry, index) => ({ ...entry, insertionOrder: index }));
          lbv2State.bulkSelectedEntryIds = new Set();
          lbv2PersistLorebook(lorebook);
          lbv2RenderEntryList();
        },
      });
    },
  }));

  lbv2OpenSheet({ title: `${selected.length} selected`, content });
}

function lbv2OpenEditorOverflowSheet() {
  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  content.appendChild(lbv2CreateSheetButton({
    label: "Duplicate Entry",
    onClick: () => {
      lbv2DuplicateActiveEntry();
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Move to Another Lorebook",
    onClick: () => {
      const entry = lbv2GetActiveEntry();
      if (!entry) return;
      lbv2MoveEntriesToLorebook([entry.id]);
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Copy as JSON",
    onClick: () => {
      const entry = lbv2GetActiveEntry();
      if (!entry) return;
      lbv2CopyText(JSON.stringify(entry, null, 2));
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "View Raw JSON",
    onClick: () => {
      const entry = lbv2GetActiveEntry();
      if (!entry) return;
      const pre = document.createElement("pre");
      pre.className = "entity-meta";
      pre.textContent = JSON.stringify(entry, null, 2);
      lbv2OpenSheet({ title: "Raw entry JSON", content: pre });
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Delete Entry",
    className: "danger",
    onClick: () => {
      lbv2DeleteActiveEntry();
      lbv2CloseSheet();
    },
  }));

  lbv2OpenSheet({ title: "Entry actions", content });
}

async function lbv2CopyText(text) {
  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(String(text || ""));
      showToast("Copied", "success");
      return;
    }
  } catch {
    // fallback below
  }

  const ta = document.createElement("textarea");
  ta.value = String(text || "");
  ta.style.position = "fixed";
  ta.style.left = "-9999px";
  document.body.appendChild(ta);
  ta.select();
  document.execCommand("copy");
  ta.remove();
  showToast("Copied", "success");
}

function lbv2DuplicateActiveEntry() {
  const lorebook = lbv2GetActiveLorebook();
  const entry = lbv2GetActiveEntry();
  if (!lorebook || !entry) return;

  const nowIso = new Date().toISOString();
  const duplicate = normalizeLorebookEntry({
    ...entry,
    id: crypto.randomUUID(),
    title: `${entry.title || "Entry"} (Copy)`,
    insertion_order: (lorebook.entries?.length || 0),
    created_at: nowIso,
    updated_at: nowIso,
  }, lorebook.entries?.length || 0);

  lorebook.entries = [...(lorebook.entries || []), duplicate]
    .map((item, index) => ({ ...item, insertionOrder: index }));

  lbv2PersistLorebook(lorebook);
  lbv2State.activeEntryId = duplicate.id;
  lbv2RenderEntryList();
  lbv2RenderEditor();
}

function lbv2DeleteActiveEntry() {
  const lorebook = lbv2GetActiveLorebook();
  const entry = lbv2GetActiveEntry();
  if (!lorebook || !entry) return;

  lbv2OpenDialog({
    title: `Delete ${entry.title || "entry"}?`,
    body: "This cannot be undone.",
    confirmLabel: "Delete",
    danger: true,
    onConfirm: () => {
      lorebook.entries = (lorebook.entries || []).filter((item) => item.id !== entry.id);
      lorebook.entries = lorebook.entries.map((item, index) => ({ ...item, insertionOrder: index }));
      lbv2PersistLorebook(lorebook);
      lbv2State.activeEntryId = "";
      lbv2NavigateTo(2);
      lbv2RenderEntryList();
      lbv2RenderEditor();
    },
  });
}

function lbv2MoveEntriesToLorebook(entryIds = []) {
  const source = lbv2GetActiveLorebook();
  const ids = (Array.isArray(entryIds) ? entryIds : []).filter(Boolean);
  if (!source || !ids.length) return;

  const otherBooks = state.lorebooks.filter((book) => book.id !== source.id);
  if (!otherBooks.length) {
    showToast("Create another lorebook first", "error");
    return;
  }

  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  otherBooks.forEach((book) => {
    content.appendChild(lbv2CreateSheetButton({
      label: book.name || "Untitled lorebook",
      sublabel: `${book.entries?.length || 0} entries`,
      onClick: () => {
        lbv2MoveEntries(ids, book.id);
        lbv2CloseSheet();
      },
    }));
  });

  lbv2OpenSheet({ title: "Move to lorebook", content });
}

function lbv2MoveEntries(entryIds = [], targetLorebookId = "") {
  const source = lbv2GetActiveLorebook();
  const target = state.lorebooks.find((book) => book.id === targetLorebookId);
  if (!source || !target || source.id === target.id) return;

  const wasBulkMode = lbv2State.bulkMode;
  const ids = new Set(entryIds);
  const moving = (source.entries || []).filter((entry) => ids.has(entry.id));
  if (!moving.length) return;

  source.entries = (source.entries || []).filter((entry) => !ids.has(entry.id));
  source.entries = source.entries.map((entry, index) => ({ ...entry, insertionOrder: index }));

  const nowIso = new Date().toISOString();
  target.entries = [
    ...(target.entries || []),
    ...moving.map((entry) => ({ ...entry, updatedAt: nowIso })),
  ].map((entry, index) => ({ ...entry, insertionOrder: index }));

  lbv2PersistLorebook(source);
  lbv2PersistLorebook(target);

  lbv2State.bulkSelectedEntryIds = new Set();
  if (wasBulkMode) {
    lbv2State.bulkMode = false;
    lbv2State.activeLorebookId = source.id;
    lbv2State.activeEntryId = "";
    lbv2NavigateTo(2);
  } else if (lbv2State.activeEntryId && ids.has(lbv2State.activeEntryId)) {
    lbv2State.activeLorebookId = target.id;
    lbv2State.activeEntryId = moving[0]?.id || "";
  }

  lbv2RenderEntryList();
  lbv2RenderEditor();
  lbv2SyncStack();
  showToast(`Moved ${moving.length} entr${moving.length === 1 ? "y" : "ies"}`, "success");
}

function lbv2OpenSettingsModal() {
  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook || !els.lbv2SettingsModal) return;

  lbv2State.suppressEditorEvents = true;

  els.lbv2SettingsNameInput && (els.lbv2SettingsNameInput.value = lorebook.name || "");
  els.lbv2SettingsDescriptionInput && (els.lbv2SettingsDescriptionInput.value = lorebook.description || "");
  els.lbv2SettingsTokenBudgetInput && (els.lbv2SettingsTokenBudgetInput.value = String(lorebook.tokenBudget || defaults.lorebookSettings.tokenBudget));
  els.lbv2SettingsDefaultScanDepthInput && (els.lbv2SettingsDefaultScanDepthInput.value = String(lorebook.defaultEntrySettings?.scanDepth ?? 10));
  els.lbv2SettingsDefaultOrderInput && (els.lbv2SettingsDefaultOrderInput.value = String(lorebook.defaultEntrySettings?.insertionOrder ?? 100));
  els.lbv2SettingsDefaultEnabledInput && (els.lbv2SettingsDefaultEnabledInput.checked = lorebook.defaultEntrySettings?.enabled !== false);
  els.lbv2SettingsDefaultCaseInput && (els.lbv2SettingsDefaultCaseInput.checked = Boolean(lorebook.defaultEntrySettings?.caseSensitive));
  els.lbv2SettingsRecursiveInput && (els.lbv2SettingsRecursiveInput.checked = lorebook.recursiveScanning !== false);
  els.lbv2SettingsRecursionDepthInput && (els.lbv2SettingsRecursionDepthInput.value = String(lorebook.maxRecursionDepth ?? 3));
  els.lbv2SettingsRecursionDepthField?.classList.toggle("hidden", !(lorebook.recursiveScanning !== false));

  const defaultPosition = lorebook.defaultEntrySettings?.insertionPosition || "after_system";
  if (els.lbv2SettingsDefaultPositionInput) els.lbv2SettingsDefaultPositionInput.value = defaultPosition;
  const defaultPositionLabel = LBV2_INJECTION_POSITIONS.find((item) => item.value === defaultPosition)?.label || "After System Prompt";
  if (els.lbv2SettingsDefaultPositionBtn) els.lbv2SettingsDefaultPositionBtn.textContent = defaultPositionLabel;

  lbv2RenderSettingsColorDots(lorebook.color || LOREBOOK_V2_COLORS[0]);

  const scope = String(lorebook.attachmentScope || "specific");
  els.lbv2Panel.querySelectorAll('input[name="lbv2AttachmentScope"]').forEach((radio) => {
    radio.checked = radio.value === scope;
  });

  if (els.lbv2AttachmentCharacterSelect) {
    const existing = state.characters || [];
    els.lbv2AttachmentCharacterSelect.innerHTML = "";
    existing.forEach((character) => {
      const option = document.createElement("option");
      option.value = character.id;
      option.textContent = character.name || "Unnamed character";
      option.selected = character.id === lorebook.attachmentCharacterId;
      els.lbv2AttachmentCharacterSelect.appendChild(option);
    });
  }

  lbv2RenderAttachmentChatChips(lorebook);
  lbv2UpdateAttachmentVisibility();

  els.lbv2SettingsModal.classList.remove("hidden");
  requestAnimationFrame(() => {
    els.lbv2SettingsModal?.classList.add("show");
    lbv2State.suppressEditorEvents = false;
  });
}

function lbv2CloseSettingsModal() {
  if (!els.lbv2SettingsModal) return;
  lbv2State.suppressEditorEvents = true;
  els.lbv2SettingsModal.classList.remove("show");
  setTimeout(() => {
    els.lbv2SettingsModal?.classList.add("hidden");
    lbv2State.suppressEditorEvents = false;
  }, 250);
}

function lbv2RenderSettingsColorDots(activeColor = LOREBOOK_V2_COLORS[0]) {
  if (!els.lbv2SettingsColorDots) return;
  els.lbv2SettingsColorDots.innerHTML = "";
  LOREBOOK_V2_COLORS.forEach((color, index) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `lbv2-color-choice${color === activeColor ? " active" : ""}`;
    btn.style.background = color;
    btn.setAttribute("aria-label", `Color ${index + 1}`);
    btn.addEventListener("click", () => {
      els.lbv2SettingsColorDots?.querySelectorAll(".lbv2-color-choice").forEach((node) => node.classList.remove("active"));
      btn.classList.add("active");
      lbv2SaveSettingsModal();
    });
    els.lbv2SettingsColorDots.appendChild(btn);
  });
}

function lbv2UpdateAttachmentVisibility() {
  const selected = els.lbv2Panel.querySelector('input[name="lbv2AttachmentScope"]:checked');
  const scope = selected?.value || "specific";
  els.lbv2AttachmentCharacterField?.classList.toggle("hidden", scope !== "character");
  els.lbv2AttachmentChatWrap?.classList.toggle("hidden", scope !== "specific");
}

function lbv2RenderAttachmentChatChips(book) {
  if (!els.lbv2AttachmentChatChipRow) return;
  els.lbv2AttachmentChatChipRow.innerHTML = "";
  const ids = Array.isArray(book?.attachedChatIds) ? book.attachedChatIds : [];

  if (!ids.length) {
    const muted = document.createElement("small");
    muted.className = "lbv2-helper";
    muted.textContent = "No chats attached";
    els.lbv2AttachmentChatChipRow.appendChild(muted);
    return;
  }

  ids.forEach((id) => {
    const chat = state.chats.find((item) => item.id === id);
    const chip = document.createElement("span");
    chip.className = "lbv2-chip";
    chip.textContent = chat?.title || "Unknown chat";
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.setAttribute("aria-label", "Remove chat");
    removeBtn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12"></path><path d="M18 6L6 18"></path></svg>`;
    removeBtn.addEventListener("click", () => {
      book.attachedChatIds = ids.filter((chatId) => chatId !== id);
      lbv2SaveSettingsModal();
      lbv2RenderAttachmentChatChips(book);
    });
    chip.appendChild(removeBtn);
    els.lbv2AttachmentChatChipRow.appendChild(chip);
  });
}

function lbv2SaveSettingsModal() {
  if (lbv2State.suppressEditorEvents) return;

  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook || !els.lbv2SettingsModal?.classList.contains("show")) return;

  lorebook.name = els.lbv2SettingsNameInput?.value?.trim() || lorebook.name || "Untitled lorebook";
  lorebook.description = els.lbv2SettingsDescriptionInput?.value?.trim() || "";

  const activeColorNode = els.lbv2SettingsColorDots?.querySelector(".lbv2-color-choice.active");
  if (activeColorNode?.style?.background) lorebook.color = activeColorNode.style.background;

  lorebook.tokenBudget = Math.min(8000, Math.max(64, Number.parseInt(els.lbv2SettingsTokenBudgetInput?.value || `${lorebook.tokenBudget || 2048}`, 10) || 2048));
  lorebook.recursiveScanning = Boolean(els.lbv2SettingsRecursiveInput?.checked ?? true);
  lorebook.maxRecursionDepth = Math.max(0, Number.parseInt(els.lbv2SettingsRecursionDepthInput?.value || `${lorebook.maxRecursionDepth || 3}`, 10) || 3);

  lorebook.defaultEntrySettings = {
    insertionPosition: els.lbv2SettingsDefaultPositionInput?.value || lorebook.defaultEntrySettings?.insertionPosition || "after_system",
    scanDepth: Math.max(0, Number.parseInt(els.lbv2SettingsDefaultScanDepthInput?.value || `${lorebook.defaultEntrySettings?.scanDepth ?? 10}`, 10) || 10),
    insertionOrder: Math.max(0, Number.parseInt(els.lbv2SettingsDefaultOrderInput?.value || `${lorebook.defaultEntrySettings?.insertionOrder ?? 100}`, 10) || 100),
    enabled: Boolean(els.lbv2SettingsDefaultEnabledInput?.checked ?? true),
    caseSensitive: Boolean(els.lbv2SettingsDefaultCaseInput?.checked ?? false),
  };

  const selectedScope = els.lbv2Panel.querySelector('input[name="lbv2AttachmentScope"]:checked');
  lorebook.attachmentScope = selectedScope?.value || "specific";
  lorebook.attachmentCharacterId = els.lbv2AttachmentCharacterSelect?.value || "";
  lorebook.attachedChatIds = normalizeStringArray(lorebook.attachedChatIds || []);

  const currentChat = getCurrentChat();
  if (currentChat?.id && lorebook.attachmentScope === "specific") {
    const inSpecific = lorebook.attachedChatIds.includes(currentChat.id);
    const set = new Set(currentChat.lorebookIds || []);
    if (inSpecific) set.add(lorebook.id);
    else set.delete(lorebook.id);
    currentChat.lorebookIds = [...set];
  }

  lbv2PersistLorebook(lorebook);
  lbv2RenderLorebookList();
  lbv2RenderEntryList();
  if (lorebook.id === lbv2State.activeLorebookId && els.lbv2EntriesBackLabel) {
    els.lbv2EntriesBackLabel.textContent = lorebook.name || "Lorebooks";
  }
}

function lbv2OpenAttachChatPicker() {
  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook) return;

  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  if (!state.chats.length) {
    const empty = document.createElement("small");
    empty.className = "lbv2-helper";
    empty.textContent = "No chats available";
    content.appendChild(empty);
  }

  state.chats.forEach((chat) => {
    const attached = Array.isArray(lorebook.attachedChatIds) && lorebook.attachedChatIds.includes(chat.id);
    const button = lbv2CreateSheetButton({
      label: chat.title || "Untitled chat",
      sublabel: attached ? "Attached" : "Tap to attach",
      onClick: () => {
        const set = new Set(Array.isArray(lorebook.attachedChatIds) ? lorebook.attachedChatIds : []);
        if (set.has(chat.id)) set.delete(chat.id);
        else set.add(chat.id);
        lorebook.attachedChatIds = [...set];

        lbv2PersistLorebook(lorebook);
        lbv2RenderAttachmentChatChips(lorebook);
        lbv2OpenAttachChatPicker();
      },
    });
    content.appendChild(button);
  });

  lbv2OpenSheet({ title: "Attach to chats", content });
}

async function lbv2ImportEntriesIntoActiveLorebook(event) {
  const lorebook = lbv2GetActiveLorebook();
  const files = Array.from(event?.target?.files || []);
  if (!lorebook || !files.length) return;

  try {
    let imported = 0;
    for (const file of files) {
      const text = await file.text();
      let parsed = tryParseLooseJson(text);
      let entries = parseLorebookEntries(parsed);

      if (!entries.length && /\.jsonl$/i.test(file.name || "")) {
        const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
        const parsedLines = lines.map((line) => {
          try { return JSON.parse(line); } catch { return null; }
        }).filter(Boolean);
        entries = parsedLines.map((entry, index) => normalizeLorebookEntry(entry, index));
      }

      if (!entries.length) continue;

      const base = lorebook.entries?.length || 0;
      const nowIso = new Date().toISOString();
      const normalized = entries.map((entry, index) => normalizeLorebookEntry({
        ...entry,
        id: crypto.randomUUID(),
        insertion_order: base + index,
        created_at: entry.createdAt || nowIso,
        updated_at: nowIso,
      }, base + index));

      lorebook.entries = [...(lorebook.entries || []), ...normalized]
        .map((item, index) => ({ ...item, insertionOrder: index }));
      imported += normalized.length;
    }

    if (imported) {
      lbv2PersistLorebook(lorebook);
      lbv2RenderEntryList();
      showToast(`Imported ${imported} entr${imported === 1 ? "y" : "ies"}`, "success");
    } else {
      showToast("No entries found in file", "error");
    }
  } catch (error) {
    showToast(`Import failed: ${error?.message || "invalid format"}`, "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function lbv2ExportActiveLorebookAsSillyTavern() {
  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook) return;

  const entries = Array.isArray(lorebook.entries) ? lorebook.entries : [];
  const worldInfoEntries = {};
  entries.forEach((entry, index) => {
    worldInfoEntries[String(index)] = {
      uid: index,
      key: entry.keys || [],
      keysecondary: entry.secondaryKeys || [],
      comment: entry.title || "",
      content: entry.content || "",
      constant: Boolean(entry.constant),
      selective: Boolean(entry.selective),
      selectiveLogic: entry.selectiveLogic || "AND",
      order: Number(entry.insertionOrder || index),
      priority: Number(entry.priority || 100),
      probability: Number(entry.probability || 100),
      disable: entry.enabled === false,
      position: entry.insertionPosition || "after_system",
    };
  });

  const payload = {
    name: lorebook.name || "Lorebook",
    description: lorebook.description || "",
    world_info: {
      entries: worldInfoEntries,
    },
  };

  downloadJson(`${slugify(lorebook.name || "lorebook")}-st.json`, payload);
  showToast("Exported as SillyTavern format", "success");
}

function lbv2GetCurrentEditorTags() {
  const entry = lbv2GetActiveEntry();
  if (!entry) return [];
  return Array.isArray(entry.tags)
    ? [...new Set(entry.tags.map((tag) => String(tag || "").trim()).filter(Boolean))]
    : [];
}

function lbv2GetCurrentEditorLinkedIds() {
  const entry = lbv2GetActiveEntry();
  if (!entry) return [];
  return Array.isArray(entry.linkedEntries)
    ? [...new Set(entry.linkedEntries.map((id) => String(id || "").trim()).filter(Boolean))]
    : [];
}

function lbv2UpdateEditorDerivedUi() {
  const primaryKeys = parseCommaSeparatedList(els.lbv2PrimaryKeysInput?.value || "");
  if (els.lbv2PrimaryKeysPreview) {
    els.lbv2PrimaryKeysPreview.innerHTML = "";
    primaryKeys.forEach((key) => {
      const chip = document.createElement("span");
      chip.className = "lbv2-chip";
      chip.textContent = key;
      els.lbv2PrimaryKeysPreview.appendChild(chip);
    });
  }

  const content = String(els.lbv2EntryContentInput?.value || "");
  const tokenCount = estimateTokenCount(content);
  if (els.lbv2ContentTokenCount) {
    els.lbv2ContentTokenCount.textContent = `${tokenCount.toLocaleString()} tk`;
  }

  const selectedStrategy = els.lbv2Panel?.querySelector('input[name="lbv2Strategy"]:checked')?.value || "normal";
  if (els.lbv2ConstantHelper) {
    els.lbv2ConstantHelper.classList.toggle("hidden", selectedStrategy !== "constant");
  }

  if (els.lbv2TriggerWarning) {
    const hasNoTrigger = selectedStrategy === "normal" && primaryKeys.length === 0;
    els.lbv2TriggerWarning.classList.toggle("hidden", !hasNoTrigger);
  }

  const position = els.lbv2InsertionPositionInput?.value || "after_system";
  els.lbv2InsertionDepthField?.classList.toggle("hidden", position !== "depth");

  const groupValue = String(els.lbv2GroupInput?.value || "").trim();
  els.lbv2GroupWeightField?.classList.toggle("hidden", !groupValue);

  lbv2UpdateSwitchAria();
}

function lbv2CollectEditorEntryDraft(existingEntry = null) {
  const activeLorebook = lbv2GetActiveLorebook();
  const current = existingEntry || lbv2GetActiveEntry();
  const nowIso = new Date().toISOString();

  const title = (els.lbv2EntryNameInput?.value || "").trim() || "Untitled Entry";
  const content = (els.lbv2EntryContentInput?.value || "").trim();
  const keys = parseCommaSeparatedList(els.lbv2PrimaryKeysInput?.value || "");
  const secondaryKeys = parseCommaSeparatedList(els.lbv2SecondaryKeysInput?.value || "");

  let activationLogic = "ANY";
  if (els.lbv2ActivationAllInput?.checked) activationLogic = "ALL";
  else if (els.lbv2ActivationAndOrInput?.checked) activationLogic = "AND_ANY_SECONDARY";
  else if (els.lbv2ActivationNotInput?.checked) activationLogic = "NOT";

  const strategy = els.lbv2Panel?.querySelector('input[name="lbv2Strategy"]:checked')?.value || "normal";
  const role = els.lbv2Panel?.querySelector('input[name="lbv2MessageRole"]:checked')?.value || "system";

  const tags = lbv2GetCurrentEditorTags();
  const linkedEntries = lbv2GetCurrentEditorLinkedIds();

  const payload = normalizeLorebookEntry({
    id: current?.id || crypto.randomUUID(),
    title,
    comment: (els.lbv2EntryCommentInput?.value || "").trim(),
    keys,
    secondary_keys: secondaryKeys,
    activation_logic: activationLogic,
    match_whole_words: Boolean(els.lbv2WholeWordsInput?.checked),
    use_regex: Boolean(els.lbv2RegexInput?.checked),
    content,
    insertion_order: Number.parseInt(els.lbv2PriorityInput?.value || `${current?.insertionOrder ?? activeLorebook?.defaultEntrySettings?.insertionOrder ?? 100}`, 10) || 100,
    insertion_position: els.lbv2InsertionPositionInput?.value || "after_system",
    insertion_depth: Math.max(0, Number.parseInt(els.lbv2InsertionDepthInput?.value || `${current?.insertionDepth ?? 4}`, 10) || 4),
    role,
    priority: Number.parseInt(els.lbv2PriorityInput?.value || `${current?.priority ?? activeLorebook?.defaultEntrySettings?.insertionOrder ?? 100}`, 10) || 100,
    case_sensitive: Boolean(els.lbv2CaseSensitiveInput?.checked),
    selective: secondaryKeys.length > 0,
    selective_logic: activationLogic === "NOT" ? "NOT" : activationLogic === "AND_ANY_SECONDARY" ? "AND" : "OR",
    scan_depth: Math.max(0, Number.parseInt(els.lbv2ScanDepthInput?.value || `${current?.scanDepth ?? activeLorebook?.defaultEntrySettings?.scanDepth ?? 10}`, 10) || 10),
    token_budget: Math.max(0, Number.parseInt(els.lbv2PerEntryTokenBudgetInput?.value || `${current?.tokenBudget ?? 0}`, 10) || 0),
    constant: strategy === "constant",
    strategy,
    prevent_recursion: Boolean(els.lbv2PreventRecursionInput?.checked),
    exclude_from_recursion: Boolean(els.lbv2ExcludeFromRecursionInput?.checked),
    probability: Math.min(100, Math.max(0, Number.parseInt(els.lbv2ActivationProbabilityInput?.value || `${current?.probability ?? 100}`, 10) || 100)),
    group: (els.lbv2GroupInput?.value || "").trim(),
    group_weight: Math.max(1, Number.parseInt(els.lbv2GroupWeightInput?.value || `${current?.groupWeight ?? 100}`, 10) || 100),
    category: (els.lbv2CategoryInput?.value || "").trim(),
    tags,
    automation_id: (els.lbv2AutomationIdInput?.value || "").trim(),
    linked_entries: linkedEntries,
    sticky: Math.max(0, Number.parseInt(els.lbv2StickyInput?.value || `${current?.sticky ?? 0}`, 10) || 0),
    cooldown: Math.max(0, Number.parseInt(els.lbv2CooldownInput?.value || `${current?.cooldown ?? 0}`, 10) || 0),
    delay: Math.max(0, Number.parseInt(els.lbv2DelayInput?.value || `${current?.delay ?? 0}`, 10) || 0),
    turn_start: Math.max(0, Number.parseInt(els.lbv2TurnStartInput?.value || `${current?.turnStart ?? 0}`, 10) || 0),
    turn_end: Number.isFinite(Number.parseInt(els.lbv2TurnEndInput?.value || "", 10))
      ? Math.max(0, Number.parseInt(els.lbv2TurnEndInput.value || "0", 10))
      : null,
    enabled: Boolean(els.lbv2EnabledInput?.checked ?? true),
    created_at: current?.createdAt || nowIso,
    updated_at: nowIso,
  });

  return payload;
}

function lbv2SaveEntryFromEditor({ silent = false } = {}) {
  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook) return;

  let existing = lbv2GetActiveEntry();
  if (!existing && lbv2State.activeEntryId) {
    existing = lbv2GetEntryById(lorebook, lbv2State.activeEntryId);
  }

  const payload = lbv2CollectEditorEntryDraft(existing);

  if (!Array.isArray(lorebook.entries)) lorebook.entries = [];
  const idx = lorebook.entries.findIndex((entry) => entry.id === payload.id);
  if (idx >= 0) {
    lorebook.entries[idx] = withEntryTimestamps({
      ...lorebook.entries[idx],
      ...payload,
    }, lorebook.entries[idx]);
  } else {
    const insertion = Number.isFinite(Number(payload.insertionOrder)) ? Number(payload.insertionOrder) : lorebook.entries.length;
    const safeIndex = Math.max(0, Math.min(lorebook.entries.length, insertion));
    lorebook.entries.splice(safeIndex, 0, withEntryTimestamps(payload, payload));
  }

  lorebook.entries = lorebook.entries.map((entry, index) => ({
    ...entry,
    insertionOrder: index,
  }));

  lbv2State.activeEntryId = payload.id;
  lbv2PersistLorebook(lorebook, { savedText: silent ? "Saved" : "Saved" });

  if (!silent) {
    lbv2RenderEntryList();
    lbv2RenderEditor();
  }
}

function lbv2RenderEntryList() {
  if (!els.lbv2EntryList) return;

  const lorebook = lbv2GetActiveLorebook();
  const entries = Array.isArray(lorebook?.entries) ? lorebook.entries.map((entry) => withEntryTimestamps(entry, entry)) : [];

  if (els.lbv2EntriesBackLabel) {
    els.lbv2EntriesBackLabel.textContent = lorebook?.name || "Lorebooks";
  }
  if (els.lbv2EntriesTitle) {
    els.lbv2EntriesTitle.textContent = lorebook?.name || "Entries";
  }

  let visible = entries.slice();
  const query = String(lbv2State.entrySearchQuery || "").trim().toLowerCase();

  if (query) {
    visible = visible.filter((entry) => {
      const haystack = [
        entry.title,
        Array.isArray(entry.keys) ? entry.keys.join(", ") : "",
        Array.isArray(entry.secondaryKeys) ? entry.secondaryKeys.join(", ") : "",
        entry.content,
      ].join(" ").toLowerCase();
      return haystack.includes(query);
    });
  }

  if (lbv2State.entryFilter === "enabled") {
    visible = visible.filter((entry) => entry.enabled !== false);
  } else if (lbv2State.entryFilter === "disabled") {
    visible = visible.filter((entry) => entry.enabled === false);
  } else if (lbv2State.entryFilter === "constant") {
    visible = visible.filter((entry) => entry.constant || entry.strategy === "constant");
  } else if (lbv2State.entryFilter === "tag" && lbv2State.entryTagFilter) {
    visible = visible.filter((entry) => Array.isArray(entry.tags) && entry.tags.includes(lbv2State.entryTagFilter));
  }

  visible = lbv2ApplyEntrySort(visible);

  if (els.lbv2FilterChipRow) {
    els.lbv2FilterChipRow.innerHTML = "";
    const chips = [];
    if (lbv2State.entryFilter && lbv2State.entryFilter !== "all") {
      chips.push({
        label: lbv2State.entryFilter === "tag" ? lbv2State.entryTagFilter : lbv2State.entryFilter,
        onRemove: () => {
          lbv2State.entryFilter = "all";
          lbv2State.entryTagFilter = "";
          lbv2RenderEntryList();
        },
      });
    }
    if (lbv2State.entrySortBy !== LOREBOOK_V2_SORT_OPTIONS.MANUAL) {
      chips.push({
        label: `Sort: ${lbv2State.entrySortBy}`,
        onRemove: () => {
          lbv2State.entrySortBy = LOREBOOK_V2_SORT_OPTIONS.MANUAL;
          lbv2RenderEntryList();
        },
      });
    }

    chips.forEach((chipData) => {
      const chip = document.createElement("span");
      chip.className = "lbv2-chip";
      chip.textContent = chipData.label;
      const x = document.createElement("button");
      x.type = "button";
      x.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12"></path><path d="M18 6L6 18"></path></svg>`;
      x.addEventListener("click", chipData.onRemove);
      chip.appendChild(x);
      els.lbv2FilterChipRow.appendChild(chip);
    });

    els.lbv2FilterChipRow.classList.toggle("hidden", chips.length === 0);
  }

  els.lbv2EntryList.innerHTML = "";

  if (!lorebook) {
    els.lbv2EntryEmptyState?.classList.remove("hidden");
    if (els.lbv2EntryCountLabel) els.lbv2EntryCountLabel.textContent = "0 entries";
    if (els.lbv2EntryTokenLabel) els.lbv2EntryTokenLabel.textContent = "0 tk total";
    return;
  }

  const totalTokens = entries.reduce((sum, entry) => sum + estimateTokenCount(entry.content || ""), 0);
  if (els.lbv2EntryCountLabel) {
    els.lbv2EntryCountLabel.textContent = `${entries.length} ${entries.length === 1 ? "entry" : "entries"}`;
  }
  if (els.lbv2EntryTokenLabel) {
    els.lbv2EntryTokenLabel.textContent = `${totalTokens.toLocaleString()} tk total`;
  }

  if (!visible.length) {
    els.lbv2EntryEmptyState?.classList.remove("hidden");
    return;
  }

  els.lbv2EntryEmptyState?.classList.add("hidden");

  visible.forEach((entry) => {
    const row = document.createElement("article");
    row.className = "lbv2-entry-row";
    row.dataset.entryId = entry.id;

    const swipeActions = document.createElement("div");
    swipeActions.className = "lbv2-swipe-actions";
    swipeActions.innerHTML = `
      <button type="button" class="lbv2-swipe-duplicate">Duplicate</button>
      <button type="button" class="lbv2-swipe-delete">Delete</button>
    `;

    const card = document.createElement("div");
    card.className = "lbv2-entry-card";
    card.innerHTML = `
      <div class="lbv2-entry-main">
        <div class="lbv2-entry-row1">
          ${lbv2State.bulkMode ? '<input type="checkbox" class="lbv2-entry-check" />' : (lbv2State.entrySortBy === LOREBOOK_V2_SORT_OPTIONS.MANUAL ? '<button type="button" class="lbv2-drag-handle" aria-label="Drag"><svg viewBox="0 0 24 24" fill="currentColor"><circle cx="8" cy="6" r="1.5"></circle><circle cx="16" cy="6" r="1.5"></circle><circle cx="8" cy="12" r="1.5"></circle><circle cx="16" cy="12" r="1.5"></circle><circle cx="8" cy="18" r="1.5"></circle><circle cx="16" cy="18" r="1.5"></circle></svg></button>' : '<button type="button" class="lbv2-status-dot-btn"></button>')}
          <h4 class="lbv2-entry-title">${escapeHtml(entry.title || "Untitled Entry")}</h4>
          <span class="lbv2-token-pill">${estimateTokenCount(entry.content || "").toLocaleString()} tk</span>
        </div>
        <div class="lbv2-entry-sub">${escapeHtml((entry.keys || []).join(", ") || "No triggers")}</div>
      </div>
    `;

    const statusBtn = card.querySelector(".lbv2-status-dot-btn");
    if (statusBtn) {
      statusBtn.classList.toggle("is-on", entry.enabled !== false);
      statusBtn.addEventListener("click", (event) => {
        event.stopPropagation();
        const lore = lbv2GetActiveLorebook();
        if (!lore) return;
        const idx = lore.entries.findIndex((item) => item.id === entry.id);
        if (idx < 0) return;
        lore.entries[idx] = {
          ...lore.entries[idx],
          enabled: lore.entries[idx].enabled === false,
          updatedAt: new Date().toISOString(),
        };
        lbv2PersistLorebook(lore);
        lbv2RenderEntryList();
      });
    }

    if (entry.constant || entry.strategy === "constant") {
      const pin = document.createElement("span");
      pin.className = "lbv2-pin";
      pin.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 17v5"></path><path d="M5 7l14 0"></path><path d="M7 7l2-4h6l2 4"></path><path d="M8 7l1 7h6l1-7"></path></svg>`;
      card.querySelector(".lbv2-entry-row1")?.appendChild(pin);
    }

    const tags = Array.isArray(entry.tags) ? entry.tags.filter(Boolean) : [];
    if (tags.length) {
      const wrap = document.createElement("div");
      wrap.className = "lbv2-chip-row";
      tags.slice(0, 4).forEach((tag) => {
        const pill = document.createElement("span");
        pill.className = "lbv2-tag-pill";
        pill.textContent = tag;
        wrap.appendChild(pill);
      });
      card.appendChild(wrap);
    }

    const checkbox = card.querySelector(".lbv2-entry-check");
    if (checkbox) {
      checkbox.checked = lbv2State.bulkSelectedEntryIds.has(entry.id);
      checkbox.addEventListener("click", (event) => {
        event.stopPropagation();
      });
    }

    const openEditor = () => {
      lbv2State.activeEntryId = entry.id;
      lbv2NavigateTo(3);
      lbv2RenderEditor();
    };

    card.addEventListener("click", () => {
      if (lbv2State.bulkMode) {
        if (lbv2State.bulkSelectedEntryIds.has(entry.id)) lbv2State.bulkSelectedEntryIds.delete(entry.id);
        else lbv2State.bulkSelectedEntryIds.add(entry.id);
        lbv2SyncStack();
        lbv2RenderEntryList();
        return;
      }
      openEditor();
    });

    const duplicateBtn = swipeActions.querySelector(".lbv2-swipe-duplicate");
    duplicateBtn?.addEventListener("click", (event) => {
      event.stopPropagation();
      const lore = lbv2GetActiveLorebook();
      if (!lore) return;
      const nowIso = new Date().toISOString();
      const duplicate = normalizeLorebookEntry({
        ...entry,
        id: crypto.randomUUID(),
        title: `${entry.title || "Entry"} (Copy)`,
        insertion_order: lore.entries.length,
        created_at: nowIso,
        updated_at: nowIso,
      }, lore.entries.length);
      lore.entries.push(duplicate);
      lore.entries = lore.entries.map((item, index) => ({ ...item, insertionOrder: index }));
      lbv2PersistLorebook(lore);
      lbv2RenderEntryList();
    });

    const deleteBtn = swipeActions.querySelector(".lbv2-swipe-delete");
    deleteBtn?.addEventListener("click", (event) => {
      event.stopPropagation();
      lbv2OpenDialog({
        title: `Delete ${entry.title || "entry"}?`,
        body: "This cannot be undone.",
        confirmLabel: "Delete",
        danger: true,
        onConfirm: () => {
          const lore = lbv2GetActiveLorebook();
          if (!lore) return;
          lore.entries = lore.entries.filter((item) => item.id !== entry.id)
            .map((item, index) => ({ ...item, insertionOrder: index }));
          lbv2PersistLorebook(lore);
          lbv2RenderEntryList();
        },
      });
    });

    let pressTimer = null;
    let startX = 0;
    let startY = 0;
    let pointerStart = null;

    const clearPress = () => {
      if (pressTimer) clearTimeout(pressTimer);
      pressTimer = null;
    };

    const openContext = () => {
      lbv2OpenEntryContextSheet(entry.id);
    };

    card.addEventListener("pointerdown", (event) => {
      if (event.pointerType === "mouse" && event.button !== 0) return;
      pointerStart = event.pointerId;
      startX = event.clientX;
      startY = event.clientY;
      clearPress();
      pressTimer = setTimeout(openContext, 520);
    });

    card.addEventListener("pointermove", (event) => {
      if (event.pointerId !== pointerStart) return;
      const dx = event.clientX - startX;
      const dy = Math.abs(event.clientY - startY);
      if (Math.abs(dx) > 10 || dy > 10) clearPress();

      if (!lbv2State.bulkMode) {
        if (dx < -40) {
          card.classList.add("is-open");
          lbv2State.expandedEntrySwipeId = entry.id;
        } else if (dx > 20) {
          card.classList.remove("is-open");
          if (lbv2State.expandedEntrySwipeId === entry.id) lbv2State.expandedEntrySwipeId = "";
        }
      }
    });

    card.addEventListener("pointerup", () => {
      clearPress();
      pointerStart = null;
      if (lbv2State.expandedEntrySwipeId && lbv2State.expandedEntrySwipeId !== entry.id) {
        card.classList.remove("is-open");
      } else if (lbv2State.expandedEntrySwipeId === entry.id) {
        card.classList.add("is-open");
      }
    });
    card.addEventListener("pointercancel", () => {
      clearPress();
      pointerStart = null;
    });

    if (lbv2State.expandedEntrySwipeId === entry.id || lbv2State.expandedEntrySwipeId === "__expand__") {
      card.classList.add("is-open");
    }

    if (lbv2State.entrySortBy === LOREBOOK_V2_SORT_OPTIONS.MANUAL) {
      const handle = card.querySelector(".lbv2-drag-handle");
      if (handle) {
        let dragTimer = null;
        let dragPointer = null;

        const clearDrag = () => {
          if (dragTimer) clearTimeout(dragTimer);
          dragTimer = null;
          dragPointer = null;
          card.classList.remove("is-drag-source");
          els.lbv2EntryList?.querySelectorAll(".lbv2-entry-card").forEach((node) => node.classList.remove("is-drop-target"));
        };

        handle.addEventListener("pointerdown", (event) => {
          event.stopPropagation();
          dragPointer = event.pointerId;
          dragTimer = setTimeout(() => {
            card.classList.add("is-drag-source");
            const move = (moveEvent) => {
              if (moveEvent.pointerId !== dragPointer) return;
              const hovered = document.elementFromPoint(moveEvent.clientX, moveEvent.clientY)?.closest?.(".lbv2-entry-row");
              els.lbv2EntryList?.querySelectorAll(".lbv2-entry-card").forEach((node) => node.classList.remove("is-drop-target"));
              hovered?.querySelector(".lbv2-entry-card")?.classList.add("is-drop-target");
            };

            const up = (upEvent) => {
              if (upEvent.pointerId !== dragPointer) return;
              document.removeEventListener("pointermove", move);
              document.removeEventListener("pointerup", up);
              document.removeEventListener("pointercancel", up);

              const hovered = document.elementFromPoint(upEvent.clientX, upEvent.clientY)?.closest?.(".lbv2-entry-row");
              const lore = lbv2GetActiveLorebook();
              if (hovered && lore) {
                const fromId = entry.id;
                const toId = hovered.dataset.entryId;
                if (fromId && toId && fromId !== toId) {
                  const arr = lore.entries || [];
                  const fromIdx = arr.findIndex((item) => item.id === fromId);
                  const toIdx = arr.findIndex((item) => item.id === toId);
                  if (fromIdx >= 0 && toIdx >= 0) {
                    const [moved] = arr.splice(fromIdx, 1);
                    arr.splice(toIdx, 0, moved);
                    lore.entries = arr.map((item, index) => ({ ...item, insertionOrder: index }));
                    lbv2PersistLorebook(lore);
                    lbv2RenderEntryList();
                  }
                }
              }

              clearDrag();
            };

            document.addEventListener("pointermove", move, { passive: false });
            document.addEventListener("pointerup", up);
            document.addEventListener("pointercancel", up);
          }, 380);
        });

        handle.addEventListener("pointerup", clearDrag);
        handle.addEventListener("pointercancel", clearDrag);
        handle.addEventListener("pointermove", (event) => {
          if (!dragTimer) return;
          if (Math.abs(event.movementX) > 8 || Math.abs(event.movementY) > 8) {
            clearDrag();
          }
        });
      }
    }

    row.append(swipeActions, card);
    els.lbv2EntryList.appendChild(row);
  });

  lbv2SyncStack();
}

function lbv2OpenEntryContextSheet(entryId) {
  const lorebook = lbv2GetActiveLorebook();
  const entry = lbv2GetEntryById(lorebook, entryId);
  if (!lorebook || !entry) return;

  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";

  content.appendChild(lbv2CreateSheetButton({
    label: entry.enabled === false ? "Enable" : "Disable",
    onClick: () => {
      const idx = lorebook.entries.findIndex((item) => item.id === entry.id);
      if (idx < 0) return;
      lorebook.entries[idx] = { ...lorebook.entries[idx], enabled: lorebook.entries[idx].enabled === false, updatedAt: new Date().toISOString() };
      lbv2PersistLorebook(lorebook);
      lbv2RenderEntryList();
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Duplicate",
    onClick: () => {
      const nowIso = new Date().toISOString();
      const duplicate = normalizeLorebookEntry({
        ...entry,
        id: crypto.randomUUID(),
        title: `${entry.title || "Entry"} (Copy)`,
        insertion_order: lorebook.entries.length,
        created_at: nowIso,
        updated_at: nowIso,
      }, lorebook.entries.length);
      lorebook.entries.push(duplicate);
      lorebook.entries = lorebook.entries.map((item, index) => ({ ...item, insertionOrder: index }));
      lbv2PersistLorebook(lorebook);
      lbv2RenderEntryList();
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Move to Another Lorebook",
    onClick: () => {
      lbv2MoveEntriesToLorebook([entry.id]);
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Copy as JSON",
    onClick: () => {
      lbv2CopyText(JSON.stringify(entry, null, 2));
      lbv2CloseSheet();
    },
  }));

  content.appendChild(lbv2CreateSheetButton({
    label: "Delete",
    className: "danger",
    onClick: () => {
      lbv2CloseSheet();
      lbv2OpenDialog({
        title: `Delete ${entry.title || "entry"}?`,
        body: "This cannot be undone.",
        confirmLabel: "Delete",
        danger: true,
        onConfirm: () => {
          lorebook.entries = lorebook.entries.filter((item) => item.id !== entry.id)
            .map((item, index) => ({ ...item, insertionOrder: index }));
          lbv2PersistLorebook(lorebook);
          lbv2RenderEntryList();
        },
      });
    },
  }));

  lbv2OpenSheet({ title: entry.title || "Entry", content });
}

function lbv2RenderEditor() {
  const lorebook = lbv2GetActiveLorebook();
  const entry = lbv2GetActiveEntry();

  if (!entry || !lorebook) {
    if (els.lbv2EditorTopTitle) els.lbv2EditorTopTitle.textContent = "Entry";
    if (els.lbv2EntryMetadata) els.lbv2EntryMetadata.textContent = "Created: — • Modified: —";
    return;
  }

  lbv2State.suppressEditorEvents = true;

  if (els.lbv2EditorTopTitle) {
    els.lbv2EditorTopTitle.textContent = entry.title || "Entry";
  }

  if (els.lbv2EntryNameInput) els.lbv2EntryNameInput.value = entry.title || "";
  if (els.lbv2EntryCommentInput) els.lbv2EntryCommentInput.value = entry.comment || "";
  if (els.lbv2PrimaryKeysInput) els.lbv2PrimaryKeysInput.value = (entry.keys || []).join(", ");
  if (els.lbv2SecondaryKeysInput) els.lbv2SecondaryKeysInput.value = (entry.secondaryKeys || []).join(", ");
  if (els.lbv2CaseSensitiveInput) els.lbv2CaseSensitiveInput.checked = Boolean(entry.caseSensitive);
  if (els.lbv2WholeWordsInput) els.lbv2WholeWordsInput.checked = Boolean(entry.matchWholeWords);
  if (els.lbv2RegexInput) els.lbv2RegexInput.checked = Boolean(entry.useRegex);

  if (els.lbv2ActivationAnyInput) els.lbv2ActivationAnyInput.checked = (entry.activationLogic || "ANY") === "ANY";
  if (els.lbv2ActivationAllInput) els.lbv2ActivationAllInput.checked = (entry.activationLogic || "ANY") === "ALL";
  if (els.lbv2ActivationAndOrInput) els.lbv2ActivationAndOrInput.checked = (entry.activationLogic || "ANY") === "AND_ANY_SECONDARY";
  if (els.lbv2ActivationNotInput) els.lbv2ActivationNotInput.checked = (entry.activationLogic || "ANY") === "NOT";

  if (els.lbv2ScanDepthInput) {
    const fallback = lorebook.defaultEntrySettings?.scanDepth ?? 10;
    els.lbv2ScanDepthInput.value = String(Number.isFinite(Number(entry.scanDepth)) ? Number(entry.scanDepth) : fallback);
  }
  if (els.lbv2PerEntryTokenBudgetInput) els.lbv2PerEntryTokenBudgetInput.value = String(Number.isFinite(Number(entry.tokenBudget)) ? Number(entry.tokenBudget) : 0);

  if (els.lbv2EntryContentInput) els.lbv2EntryContentInput.value = entry.content || "";

  if (els.lbv2InsertionPositionInput) els.lbv2InsertionPositionInput.value = entry.insertionPosition || "after_system";
  if (els.lbv2InsertionPositionBtn) {
    const label = LBV2_INJECTION_POSITIONS.find((item) => item.value === (entry.insertionPosition || "after_system"))?.label || "After System Prompt";
    els.lbv2InsertionPositionBtn.textContent = label;
  }
  if (els.lbv2InsertionDepthInput) els.lbv2InsertionDepthInput.value = String(Number.isFinite(Number(entry.insertionDepth)) ? Number(entry.insertionDepth) : 4);
  if (els.lbv2PriorityInput) els.lbv2PriorityInput.value = String(Number.isFinite(Number(entry.priority)) ? Number(entry.priority) : 100);

  const role = entry.role || "system";
  els.lbv2Panel?.querySelectorAll('input[name="lbv2MessageRole"]').forEach((radio) => {
    radio.checked = radio.value === role;
  });

  if (els.lbv2EnabledInput) els.lbv2EnabledInput.checked = entry.enabled !== false;

  const strategy = entry.strategy || (entry.constant ? "constant" : "normal");
  els.lbv2Panel?.querySelectorAll('input[name="lbv2Strategy"]').forEach((radio) => {
    radio.checked = radio.value === strategy;
  });

  if (els.lbv2PreventRecursionInput) els.lbv2PreventRecursionInput.checked = Boolean(entry.preventRecursion);
  if (els.lbv2ExcludeFromRecursionInput) els.lbv2ExcludeFromRecursionInput.checked = Boolean(entry.excludeFromRecursion);

  const probability = Math.min(100, Math.max(0, Number.parseInt(entry.probability ?? 100, 10) || 100));
  if (els.lbv2ActivationProbabilitySlider) els.lbv2ActivationProbabilitySlider.value = String(probability);
  if (els.lbv2ActivationProbabilityInput) els.lbv2ActivationProbabilityInput.value = String(probability);

  if (els.lbv2GroupInput) els.lbv2GroupInput.value = entry.group || "";
  if (els.lbv2GroupWeightInput) els.lbv2GroupWeightInput.value = String(Number.isFinite(Number(entry.groupWeight)) ? Number(entry.groupWeight) : 100);

  if (els.lbv2CategoryInput) els.lbv2CategoryInput.value = entry.category || "";
  if (els.lbv2CategoryPickerBtn) {
    els.lbv2CategoryPickerBtn.textContent = entry.category || "Uncategorized";
  }

  if (els.lbv2AutomationIdInput) {
    const auto = entry.automationId || `entry_${slugify(entry.title || "entry")}_${String((entry.insertionOrder || 0) + 1).padStart(3, "0")}`;
    els.lbv2AutomationIdInput.value = auto;
  }

  if (els.lbv2StickyInput) els.lbv2StickyInput.value = String(Number.isFinite(Number(entry.sticky)) ? Number(entry.sticky) : 0);
  if (els.lbv2CooldownInput) els.lbv2CooldownInput.value = String(Number.isFinite(Number(entry.cooldown)) ? Number(entry.cooldown) : 0);
  if (els.lbv2DelayInput) els.lbv2DelayInput.value = String(Number.isFinite(Number(entry.delay)) ? Number(entry.delay) : 0);
  if (els.lbv2TurnStartInput) els.lbv2TurnStartInput.value = String(Number.isFinite(Number(entry.turnStart)) ? Number(entry.turnStart) : 0);
  if (els.lbv2TurnEndInput) els.lbv2TurnEndInput.value = Number.isFinite(Number(entry.turnEnd)) ? String(Number(entry.turnEnd)) : "";

  lbv2RenderTagList(Array.isArray(entry.tags) ? entry.tags : []);
  lbv2RenderLinkedEntries(Array.isArray(entry.linkedEntries) ? entry.linkedEntries : []);

  if (els.lbv2EntryMetadata) {
    els.lbv2EntryMetadata.textContent = `Created: ${formatDateShort(entry.createdAt)} • Modified: ${formatRelativeTime(entry.updatedAt)}`;
  }

  lbv2UpdateEditorDerivedUi();
  lbv2State.suppressEditorEvents = false;
}

function lbv2RenderTagList(tags = []) {
  if (!els.lbv2TagList) return;
  const unique = [...new Set((Array.isArray(tags) ? tags : []).map((tag) => String(tag || "").trim()).filter(Boolean))];
  els.lbv2TagList.innerHTML = "";

  unique.forEach((tag) => {
    const chip = document.createElement("span");
    chip.className = "lbv2-chip";
    chip.textContent = tag;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.setAttribute("aria-label", `Remove ${tag}`);
    remove.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12"></path><path d="M18 6L6 18"></path></svg>`;
    remove.addEventListener("click", () => {
      const entry = lbv2GetActiveEntry();
      if (!entry) return;
      entry.tags = unique.filter((item) => item !== tag);
      lbv2RenderTagList(entry.tags);
      lbv2ScheduleAutosave();
    });
    chip.appendChild(remove);
    els.lbv2TagList.appendChild(chip);
  });

  const entry = lbv2GetActiveEntry();
  if (entry) entry.tags = unique;
}

function lbv2AddTagFromInput() {
  const input = els.lbv2TagInput;
  const entry = lbv2GetActiveEntry();
  if (!input || !entry) return;

  const value = String(input.value || "").trim().replace(/,+$/, "");
  if (!value) {
    lbv2RenderTagSuggestions();
    return;
  }

  const tags = [...new Set([...(Array.isArray(entry.tags) ? entry.tags : []), value])];
  entry.tags = tags;
  input.value = "";
  lbv2RenderTagList(tags);
  lbv2RenderTagSuggestions();
  lbv2ScheduleAutosave();
}

function lbv2RenderTagSuggestions() {
  if (!els.lbv2TagSuggestionList || !els.lbv2TagInput) return;
  const lorebook = lbv2GetActiveLorebook();
  const query = String(els.lbv2TagInput.value || "").trim().toLowerCase();

  const allTags = [...new Set((lorebook?.entries || [])
    .flatMap((entry) => Array.isArray(entry.tags) ? entry.tags : [])
    .map((tag) => String(tag || "").trim())
    .filter(Boolean))];

  const filtered = query
    ? allTags.filter((tag) => tag.toLowerCase().includes(query))
    : allTags.slice(0, 6);

  els.lbv2TagSuggestionList.innerHTML = "";
  if (!filtered.length || !query) {
    els.lbv2TagSuggestionList.classList.add("hidden");
    return;
  }

  filtered.slice(0, 8).forEach((tag) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "lbv2-suggestion-item";
    btn.textContent = tag;
    btn.addEventListener("click", () => {
      const entry = lbv2GetActiveEntry();
      if (!entry) return;
      entry.tags = [...new Set([...(entry.tags || []), tag])];
      if (els.lbv2TagInput) els.lbv2TagInput.value = "";
      lbv2RenderTagList(entry.tags);
      lbv2RenderTagSuggestions();
      lbv2ScheduleAutosave();
    });
    els.lbv2TagSuggestionList.appendChild(btn);
  });

  els.lbv2TagSuggestionList.classList.remove("hidden");
}

function lbv2RenderLinkedEntries(linkedIds = []) {
  if (!els.lbv2LinkedEntryList) return;
  const lorebook = lbv2GetActiveLorebook();
  const ids = [...new Set((Array.isArray(linkedIds) ? linkedIds : []).map((id) => String(id || "").trim()).filter(Boolean))];
  els.lbv2LinkedEntryList.innerHTML = "";

  ids.forEach((id) => {
    const linked = lbv2GetEntryById(lorebook, id);
    if (!linked) return;
    const chip = document.createElement("span");
    chip.className = "lbv2-chip";
    chip.textContent = linked.title || "Entry";
    const remove = document.createElement("button");
    remove.type = "button";
    remove.setAttribute("aria-label", "Remove linked entry");
    remove.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12"></path><path d="M18 6L6 18"></path></svg>`;
    remove.addEventListener("click", () => {
      const entry = lbv2GetActiveEntry();
      if (!entry) return;
      entry.linkedEntries = ids.filter((entryId) => entryId !== id);
      lbv2RenderLinkedEntries(entry.linkedEntries);
      lbv2ScheduleAutosave();
    });
    chip.appendChild(remove);
    els.lbv2LinkedEntryList.appendChild(chip);
  });

  const entry = lbv2GetActiveEntry();
  if (entry) entry.linkedEntries = ids;
}

function lbv2OpenLinkEntryPicker() {
  const lorebook = lbv2GetActiveLorebook();
  const entry = lbv2GetActiveEntry();
  if (!lorebook || !entry) return;

  const content = document.createElement("div");
  content.className = "lbv2-sheet-content";
  const selected = new Set(Array.isArray(entry.linkedEntries) ? entry.linkedEntries : []);

  (lorebook.entries || []).forEach((item) => {
    if (item.id === entry.id) return;
    content.appendChild(lbv2CreateSheetButton({
      label: item.title || "Untitled entry",
      sublabel: selected.has(item.id) ? "Linked" : "",
      onClick: () => {
        if (selected.has(item.id)) selected.delete(item.id);
        else selected.add(item.id);
        entry.linkedEntries = [...selected];
        lbv2RenderLinkedEntries(entry.linkedEntries);
        lbv2ScheduleAutosave();
        lbv2OpenLinkEntryPicker();
      },
    }));
  });

  lbv2OpenSheet({ title: "Link Entries", content });
}

function lbv2OpenPickerSheet(kind = "") {
  if (kind === "insertion-position") {
    const content = document.createElement("div");
    content.className = "lbv2-sheet-content";
    LBV2_INJECTION_POSITIONS.forEach((option) => {
      content.appendChild(lbv2CreateSheetButton({
        label: option.label,
        sublabel: els.lbv2InsertionPositionInput?.value === option.value ? "Selected" : "",
        onClick: () => {
          if (els.lbv2InsertionPositionInput) els.lbv2InsertionPositionInput.value = option.value;
          if (els.lbv2InsertionPositionBtn) els.lbv2InsertionPositionBtn.textContent = option.label;
          lbv2UpdateEditorDerivedUi();
          lbv2ScheduleAutosave();
          lbv2CloseSheet();
        },
      }));
    });
    lbv2OpenSheet({ title: "Insertion Position", content });
    return;
  }

  if (kind === "default-position") {
    const content = document.createElement("div");
    content.className = "lbv2-sheet-content";
    LBV2_INJECTION_POSITIONS.forEach((option) => {
      content.appendChild(lbv2CreateSheetButton({
        label: option.label,
        sublabel: els.lbv2SettingsDefaultPositionInput?.value === option.value ? "Selected" : "",
        onClick: () => {
          if (els.lbv2SettingsDefaultPositionInput) els.lbv2SettingsDefaultPositionInput.value = option.value;
          if (els.lbv2SettingsDefaultPositionBtn) els.lbv2SettingsDefaultPositionBtn.textContent = option.label;
          lbv2SaveSettingsModal();
          lbv2CloseSheet();
        },
      }));
    });
    lbv2OpenSheet({ title: "Default Position", content });
    return;
  }

  if (kind === "category") {
    const lorebook = lbv2GetActiveLorebook();
    const entry = lbv2GetActiveEntry();
    if (!lorebook || !entry) return;

    const categories = [...new Set((lorebook.entries || [])
      .map((item) => String(item.category || "").trim())
      .filter(Boolean))];

    const content = document.createElement("div");
    content.className = "lbv2-sheet-content";

    content.appendChild(lbv2CreateSheetButton({
      label: "Uncategorized",
      sublabel: !entry.category ? "Selected" : "",
      onClick: () => {
        if (els.lbv2CategoryInput) els.lbv2CategoryInput.value = "";
        if (els.lbv2CategoryPickerBtn) els.lbv2CategoryPickerBtn.textContent = "Uncategorized";
        lbv2ScheduleAutosave();
        lbv2CloseSheet();
      },
    }));

    categories.forEach((category) => {
      content.appendChild(lbv2CreateSheetButton({
        label: category,
        sublabel: entry.category === category ? "Selected" : "",
        onClick: () => {
          if (els.lbv2CategoryInput) els.lbv2CategoryInput.value = category;
          if (els.lbv2CategoryPickerBtn) els.lbv2CategoryPickerBtn.textContent = category;
          lbv2ScheduleAutosave();
          lbv2CloseSheet();
        },
      }));
    });

    content.appendChild(lbv2CreateSheetButton({
      label: "+ New Category",
      onClick: () => {
        const form = document.createElement("div");
        form.className = "lbv2-sheet-form";
        const input = document.createElement("input");
        input.type = "text";
        input.placeholder = "Category name";
        const create = document.createElement("button");
        create.type = "button";
        create.className = "lbv2-primary-btn";
        create.textContent = "Create";
        create.addEventListener("click", () => {
          const value = input.value.trim();
          if (!value) return;
          if (els.lbv2CategoryInput) els.lbv2CategoryInput.value = value;
          if (els.lbv2CategoryPickerBtn) els.lbv2CategoryPickerBtn.textContent = value;
          lbv2ScheduleAutosave();
          lbv2CloseSheet();
        });
        form.append(input, create);
        lbv2OpenSheet({ title: "New Category", content: form });
        setTimeout(() => input.focus(), 40);
      },
    }));

    lbv2OpenSheet({ title: "Category", content });
  }
}

function lbv2CreateEntryAndOpenEditor() {
  const lorebook = lbv2GetActiveLorebook();
  if (!lorebook) return;

  if (!Array.isArray(lorebook.entries)) lorebook.entries = [];
  const entry = lbv2CreateEmptyEntry(lorebook);
  lorebook.entries.push(entry);
  lorebook.entries = lorebook.entries.map((item, index) => ({ ...item, insertionOrder: index }));

  lbv2PersistLorebook(lorebook);
  lbv2State.activeEntryId = entry.id;
  lbv2NavigateTo(3);
  lbv2RenderEntryList();
  lbv2RenderEditor();
  lbv2ShowSaveIndicator("Saved");
  setTimeout(() => els.lbv2EntryNameInput?.focus(), 40);
}

function renderLorebooks() {
  state.lorebooks = (Array.isArray(state.lorebooks) ? state.lorebooks : []).map((book) => normalizeLorebookRecord(book));

  if (LOREBOOK_V2_ENABLED && els.lbv2Panel && els.lbv2LorebookList) {
    renderLorebooksV2();
    return;
  }

  if (!els.lorebookList) return;

  const query = els.lorebookSearchInput?.value?.trim().toLowerCase() || "";
  const attachedIds = new Set(state.chats.flatMap((chat) => Array.isArray(chat?.lorebookIds) ? chat.lorebookIds : []));

  const lorebooks = state.lorebooks.filter((book) => {
    if (!query) return true;
    const haystack = `${book.name} ${book.description}`.toLowerCase();
    return haystack.includes(query);
  });

  els.lorebookList.innerHTML = "";
  if (!lorebooks.length) {
    const empty = document.createElement("article");
    empty.className = "placeholder-card";
    empty.innerHTML = `<p>No lorebooks yet. Import one or create a new lorebook.</p>`;
    els.lorebookList.appendChild(empty);
    return;
  }

  lorebooks.forEach((book) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-actions">
        <div class="entity-card-header" style="width:100%; margin-bottom:0;">
          <h3>${escapeHtml(book.name || "Untitled lorebook")}</h3>
        </div>
      </div>
      <div class="entity-meta">${escapeHtml(book.description || "No description")}</div>
      <div class="entity-meta">Entries ${book.entries.length} • Default scan ${book.scanDepth} • Default budget ${book.tokenBudget}</div>
      <div class="entity-meta">Attached in ${attachedIds.has(book.id) ? state.chats.filter((chat) => Array.isArray(chat?.lorebookIds) && chat.lorebookIds.includes(book.id)).length : 0} chat(s)</div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="edit" type="button">Edit</button>
        <button class="secondary" data-action="export" type="button">Export</button>
        <button class="danger" data-action="remove" type="button">Delete</button>
      </div>
    `;

    const actions = card.querySelector(".character-inline-actions");
    card.querySelector('[data-action="toggle-actions"]')?.addEventListener("click", () => {
      actions?.classList.toggle("hidden");
    });
    card.querySelector('[data-action="edit"]')?.addEventListener("click", () => showLorebookForm(book));
    card.querySelector('[data-action="export"]')?.addEventListener("click", () => exportLorebook(book.id));
    card.querySelector('[data-action="remove"]')?.addEventListener("click", () => removeLorebook(book.id));

    els.lorebookList.appendChild(card);
  });
}

function normalizePresetPromptList(prompts = []) {
  const toBool = (value, fallback = false) => {
    if (value === undefined || value === null) return fallback;
    if (typeof value === "boolean") return value;
    if (typeof value === "string") {
      const normalized = value.trim().toLowerCase();
      if (["true", "1", "yes", "on"].includes(normalized)) return true;
      if (["false", "0", "no", "off"].includes(normalized)) return false;
    }
    return Boolean(value);
  };

  const toPosition = (value) => {
    if (value === "depth") return 1;
    if (value === "ordered") return 0;
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : 0;
  };

  return prompts
    .map((prompt, index) => {
      const identifier = (prompt?.identifier || `prompt-${index + 1}`).trim();
      return {
        identifier,
        name: prompt?.name || identifier,
        role: prompt?.role || "system",
        content: prompt?.content || "",
        marker: toBool(prompt?.marker, false),
        system_prompt: toBool(prompt?.system_prompt, true),
        enabled: toBool(prompt?.enabled, true),
        injection_position: toPosition(prompt?.injection_position),
        injection_depth: Number.isFinite(Number(prompt?.injection_depth)) ? Number(prompt?.injection_depth) : 4,
        forbid_overrides: toBool(prompt?.forbid_overrides, false),
      };
    })
    .filter((prompt) => prompt.identifier);
}

function normalizePreset(preset) {
  return {
    id: preset.id || crypto.randomUUID(),
    name: preset.name || "Imported preset",
    systemPrompt: preset.systemPrompt || preset.content || defaults.systemPrompt,
    useSystemPrompt: Boolean(preset.useSystemPrompt ?? preset.use_sysprompt ?? true),
    namesBehavior: preset.namesBehavior || preset.names_behavior || "",
    temperature: Number.isFinite(Number(preset.temperature)) ? Number(preset.temperature) : 0.7,
    topP: Number.isFinite(Number(preset.topP ?? preset.top_p)) ? Number(preset.topP ?? preset.top_p) : 1,
    maxTokens: Number.isFinite(Number(preset.maxTokens ?? preset.openai_max_tokens)) ? Number(preset.maxTokens ?? preset.openai_max_tokens) : 300,
    stream: Boolean(preset.stream ?? preset.streaming ?? preset.stream_openai ?? false),
    contextTemplate: preset.contextTemplate || preset.story_string || defaults.presets[0].contextTemplate,
    chatStart: preset.chatStart || preset.chat_start || defaults.presets[0].chatStart,
    exampleSeparator: preset.exampleSeparator || preset.example_separator || defaults.presets[0].exampleSeparator,
    scenarioFormat: preset.scenarioFormat || preset.scenario_format || "{{scenario}}",
    personalityFormat: preset.personalityFormat || preset.personality_format || "{{personality}}",
    prompts: normalizePresetPromptList(Array.isArray(preset.prompts) ? preset.prompts : []),
    promptOrder: Array.isArray(preset.promptOrder) ? preset.promptOrder : Array.isArray(preset.prompt_order) ? preset.prompt_order : [],
    instruct: {
      enabled: Boolean(preset.instruct?.enabled ?? (preset.source === "sillytavern-instruct")),
      inputSequence: preset.instruct?.inputSequence || preset.input_sequence || defaults.presets[0].instruct.inputSequence,
      outputSequence: preset.instruct?.outputSequence || preset.output_sequence || defaults.presets[0].instruct.outputSequence,
      systemSequence: preset.instruct?.systemSequence || preset.system_sequence || defaults.presets[0].instruct.systemSequence,
      stopSequence: preset.instruct?.stopSequence || preset.stop_sequence || defaults.presets[0].instruct.stopSequence,
      wrap: Boolean(preset.instruct?.wrap ?? preset.wrap ?? false),
      macro: Boolean(preset.instruct?.macro ?? preset.macro ?? false),
      namesBehavior: preset.instruct?.namesBehavior || preset.names_behavior || "",
      inputSuffix: preset.instruct?.inputSuffix || preset.input_suffix || "",
      outputSuffix: preset.instruct?.outputSuffix || preset.output_suffix || "",
      systemSuffix: preset.instruct?.systemSuffix || preset.system_suffix || "",
      firstInputSequence: preset.instruct?.firstInputSequence || preset.first_input_sequence || "",
      firstOutputSequence: preset.instruct?.firstOutputSequence || preset.first_output_sequence || "",
      lastInputSequence: preset.instruct?.lastInputSequence || preset.last_input_sequence || "",
      lastOutputSequence: preset.instruct?.lastOutputSequence || preset.last_output_sequence || "",
      systemSameAsUser: Boolean(preset.instruct?.systemSameAsUser ?? preset.system_same_as_user ?? false),
      storyStringPrefix: preset.instruct?.storyStringPrefix || preset.story_string_prefix || "",
      storyStringSuffix: preset.instruct?.storyStringSuffix || preset.story_string_suffix || "",
    },
    source: preset.source || "maitavern",
    raw: preset.raw || null,
  };
}

function normalizeEndpoint(endpoint) {
  const trimmed = endpoint.trim().replace(/\/+$/, "");
  if (/\/chat\/completions$/i.test(trimmed)) return trimmed;
  return `${trimmed}/chat/completions`;
}

async function requestChatCompletion({ headers, payload, signal }) {
  const requestUrl = normalizeEndpoint(state.endpoint);
  const response = await fetch(requestUrl, {
    method: "POST",
    headers,
    body: JSON.stringify(payload),
    signal,
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${extractErrorMessage(text) || response.statusText}`);
  }

  const contentType = response.headers.get("content-type") || "";
  return parseApiJsonResponse(text, requestUrl, "chat completion", contentType);
}

function shouldRetryWithCompatPayload(data) {
  const choice = Array.isArray(data?.choices) ? data.choices[0] : null;
  return Boolean(choice && choice.message && choice.message.content == null);
}

function extractAssistantText(data) {
  const choice = Array.isArray(data?.choices) ? data.choices[0] : null;
  const messageContent = choice?.message?.content;

  if (typeof messageContent === "string" && messageContent.trim()) {
    return messageContent;
  }

  if (Array.isArray(messageContent)) {
    const joined = messageContent
      .map((item) => {
        if (typeof item === "string") return item;
        if (typeof item?.text === "string") return item.text;
        if (typeof item?.content === "string") return item.content;
        return "";
      })
      .filter(Boolean)
      .join("\n");
    if (joined.trim()) return joined;
  }

  if (typeof choice?.text === "string" && choice.text.trim()) {
    return choice.text;
  }

  if (typeof data?.response === "string" && data.response.trim()) return data.response;
  if (typeof data?.message === "string" && data.message.trim()) return data.message;
  if (typeof data?.content === "string" && data.content.trim()) return data.content;
  if (typeof data?.output_text === "string" && data.output_text.trim()) return data.output_text;

  if (Array.isArray(data?.output)) {
    const outputText = data.output
      .flatMap((item) => item.content || [])
      .map((item) => {
        if (typeof item === "string") return item;
        if (item.type === "output_text") return item.text;
        if (typeof item.text === "string") return item.text;
        return "";
      })
      .filter(Boolean)
      .join("\n");
    if (outputText.trim()) return outputText;
  }

  return "";
}

async function fetchModelsIntoSettings() {
  setInlineStatus(els.fetchModelsStatus, "Fetching models...");
  try {
    const models = await fetchModelsForProvider({
      endpoint: els.endpointInput.value.trim(),
      apiKey: els.apiKeyInput.value.trim(),
      authHeader: els.authHeaderInput.value.trim() || defaults.authHeader,
      authPrefix: els.authPrefixInput.value,
      extraHeadersRaw: els.extraHeadersInput.value.trim(),
    });
    applyFetchedModels(models);
    setInlineStatus(els.fetchModelsStatus, `Loaded ${models.length} model${models.length === 1 ? "" : "s"}`);
    setStatus("Models fetched");
  } catch (error) {
    setInlineStatus(els.fetchModelsStatus, error.message);
    setStatus("Model fetch failed");
  }
}

async function fetchModelsIntoProviderForm() {
  setInlineStatus(els.providerFetchModelsStatus, "Fetching models...");
  try {
    const models = await fetchModelsForProvider({
      endpoint: els.providerUrlInput.value.trim(),
      apiKey: els.providerApiKeyInput.value.trim(),
      authHeader: state.authHeader || defaults.authHeader,
      authPrefix: state.authPrefix,
      extraHeadersRaw: state.extraHeaders,
    });
    applyFetchedModels(models);
    setInlineStatus(els.providerFetchModelsStatus, `Loaded ${models.length} model${models.length === 1 ? "" : "s"}`);
    setStatus("Models fetched");
  } catch (error) {
    setInlineStatus(els.providerFetchModelsStatus, error.message);
    setStatus("Model fetch failed");
  }
}

async function fetchModelsForProvider({ endpoint, apiKey, authHeader, authPrefix, extraHeadersRaw }) {
  if (!endpoint) throw new Error("Add an endpoint first");
  if (!apiKey) throw new Error("Add an API key first");

  const headers = {
    ...parseJsonObject(extraHeadersRaw, "Extra headers"),
  };

  if (apiKey && authHeader) {
    headers[authHeader] = `${authPrefix || ""}${apiKey}`;
  }

  const modelsUrl = normalizeModelsEndpoint(endpoint);
  const response = await fetch(modelsUrl, {
    method: "GET",
    headers,
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${extractErrorMessage(text) || response.statusText}`);
  }

  const contentType = response.headers.get("content-type") || "";
  const data = parseApiJsonResponse(text, modelsUrl, "models", contentType);
  const models = extractModels(data);
  if (!models.length) {
    throw new Error("No models returned");
  }
  return models;
}

function normalizeModelsEndpoint(endpoint) {
  const trimmed = endpoint.trim().replace(/\/+$/, "");
  if (/\/models$/i.test(trimmed)) return trimmed;
  if (/\/v1\/chat\/completions$/i.test(trimmed)) {
    return trimmed.replace(/\/chat\/completions$/i, "/models");
  }
  if (/\/chat\/completions$/i.test(trimmed)) {
    return trimmed.replace(/\/chat\/completions$/i, "/models");
  }
  if (/\/v1$/i.test(trimmed)) return `${trimmed}/models`;
  return `${trimmed}/models`;
}

function extractModels(data) {
  const list = Array.isArray(data?.data)
    ? data.data.map((item) => item?.id || item?.name).filter(Boolean)
    : Array.isArray(data?.models)
      ? data.models.map((item) => typeof item === "string" ? item : item?.id || item?.name).filter(Boolean)
      : [];

  return [...new Set(list.map((item) => String(item).trim()).filter(Boolean))];
}

function applyFetchedModels(models) {
  state.models = models;
  state.selectedModel = models.includes(state.selectedModel) ? state.selectedModel : models[0];
  state.hasFetchedModels = true;
  els.modelsInput.value = models.join("\n");
  persistState();
  renderModelOptions();
}

async function maybeAutoFetchModels(force = false) {
  if (!force && state.hasFetchedModels) return;
  try {
    const models = await fetchModelsForProvider({
      endpoint: state.endpoint,
      apiKey: state.apiKey,
      authHeader: state.authHeader || defaults.authHeader,
      authPrefix: state.authPrefix,
      extraHeadersRaw: state.extraHeaders,
    });
    applyFetchedModels(models);
    setInlineStatus(els.fetchModelsStatus, `Loaded ${models.length} model${models.length === 1 ? "" : "s"}`);
    setStatus("Models fetched");
  } catch (error) {
    setInlineStatus(els.fetchModelsStatus, error.message);
  }
}

function setInlineStatus(element, text) {
  if (!element) return;
  element.textContent = text;
}

function setDataStatus(text) {
  if (!els.dataStatus) return;
  els.dataStatus.textContent = text;
}

function parseJsonObject(raw, label) {
  if (!raw) return {};
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error(`${label} must be valid JSON`);
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error(`${label} must be a JSON object`);
  }
  return parsed;
}

async function pickAndLoadUserFile() {
  const { parsed } = await readUserFileFromDisk();
  const importedState = parsed?.state && typeof parsed.state === "object" ? parsed.state : parsed;
  state = {
    ...structuredClone(defaults),
    ...(importedState && typeof importedState === "object" ? importedState : {}),
  };
}

function extractErrorMessage(text) {
  const normalized = String(text || "").trim();
  if (!normalized) return "";
  if (/^<!doctype html/i.test(normalized) || /^<html/i.test(normalized)) {
    return "The server returned HTML instead of JSON. Check the endpoint URL.";
  }
  try {
    const parsed = JSON.parse(normalized);
    return parsed?.error?.message || parsed?.message || normalized;
  } catch {
    return normalized;
  }
}

function parseStreamLikeJsonResponse(text) {
  const source = String(text || "");
  if (!source) return null;

  const chunks = [];
  for (const rawLine of source.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line.startsWith("data:")) continue;

    const payload = line.slice(5).trim();
    if (!payload || payload === "[DONE]") continue;

    try {
      chunks.push(JSON.parse(payload));
    } catch {
      // ignore malformed chunks; we'll fail only if no valid JSON chunks exist.
    }
  }

  if (!chunks.length) return null;

  const mergedText = chunks
    .map((chunk) => {
      const choice = Array.isArray(chunk?.choices) ? chunk.choices[0] : null;
      const delta = choice?.delta;

      if (typeof delta?.content === "string") return delta.content;
      if (Array.isArray(delta?.content)) {
        return delta.content
          .map((item) => {
            if (typeof item === "string") return item;
            if (typeof item?.text === "string") return item.text;
            if (typeof item?.content === "string") return item.content;
            return "";
          })
          .filter(Boolean)
          .join("");
      }

      if (typeof choice?.message?.content === "string") return choice.message.content;
      if (typeof choice?.text === "string") return choice.text;
      if (typeof chunk?.response === "string") return chunk.response;
      if (typeof chunk?.message === "string") return chunk.message;
      if (typeof chunk?.content === "string") return chunk.content;
      return "";
    })
    .join("");

  const lastChunk = chunks[chunks.length - 1];
  if (!mergedText.trim()) return lastChunk;

  const lastChoice = Array.isArray(lastChunk?.choices) ? (lastChunk.choices[0] || {}) : {};
  return {
    ...lastChunk,
    choices: [
      {
        ...lastChoice,
        message: {
          ...(lastChoice?.message && typeof lastChoice.message === "object" ? lastChoice.message : {}),
          content: mergedText,
        },
      },
    ],
  };
}

function parseApiJsonResponse(text, requestUrl, label, contentType = "") {
  const normalized = String(text || "").trim();
  if (!normalized) {
    throw new Error(`Empty ${label} response from ${requestUrl}`);
  }
  if (/^<!doctype html/i.test(normalized) || /^<html/i.test(normalized) || normalized.startsWith("<")) {
    throw new Error(`The ${label} endpoint returned HTML instead of JSON. Check the URL: ${requestUrl}`);
  }

  const normalizedContentType = String(contentType || "").toLowerCase();
  if (normalizedContentType.includes("text/event-stream")) {
    const parsedStream = parseStreamLikeJsonResponse(normalized);
    if (parsedStream) return parsedStream;
    throw new Error(`The ${label} endpoint returned an unreadable event stream. Check the URL: ${requestUrl}`);
  }

  try {
    return JSON.parse(normalized);
  } catch {
    const parsedStream = parseStreamLikeJsonResponse(normalized);
    if (parsedStream) return parsedStream;
    throw new Error(`The ${label} endpoint returned non-JSON data. Check the URL: ${requestUrl}`);
  }
}

function renderModelOptions() {
  els.modelSelect.innerHTML = "";
  for (const model of state.models) {
    const option = document.createElement("option");
    option.value = model;
    option.textContent = model;
    option.selected = model === state.selectedModel;
    els.modelSelect.appendChild(option);
  }
}

function renderMessages() {
  updateComposerOffset();
  const previousTop = els.messages.scrollTop;
  const shouldStickToBottom = Math.abs((els.messages.scrollHeight - els.messages.clientHeight) - previousTop) < 80;
  els.messages.innerHTML = "";

  if (expandedMessageMenuIndex >= state.messages.length) {
    expandedMessageMenuIndex = -1;
  }
  if (messageInlineEditIndex >= state.messages.length) {
    messageInlineEditIndex = -1;
  }

  for (let index = 0; index < state.messages.length; index += 1) {
    const msg = state.messages[index];
    const row = els.messageTemplate.content.firstElementChild.cloneNode(true);
    row.classList.add(msg.role);
    if (msg.pending) row.classList.add("pending");
    const avatar = row.querySelector(".message-avatar");
    const bubble = row.querySelector(".message");
    bubble.classList.add(msg.role);

    const style = getActiveChatStyle();
    row.dataset.chatStyle = style;
    const roleEl = row.querySelector(".message-role");
    const roleLabel = getMessageDisplayRole(msg);
    roleEl.textContent = roleLabel;
    roleEl.classList.toggle("hidden", !roleLabel);

    const showAvatar = shouldShowMessageAvatar(msg);
    row.dataset.hasAvatar = showAvatar ? "1" : "0";
    avatar.classList.toggle("hidden", !showAvatar);
    avatar.classList.remove("shape-cutout", "shape-circle", "shape-square", "shape-rectangle");
    const avatarShape = getMessageAvatarShape(msg);
    avatar.classList.add(`shape-${avatarShape}`);
    const avatarSize = getMessageAvatarSize(msg);
    const avatarWidth = avatarShape === "rectangle" ? Math.round(avatarSize * 1.28) : avatarSize;
    avatar.style.width = `${avatarWidth}px`;
    avatar.style.height = `${avatarSize}px`;
    row.style.setProperty("--message-avatar-size", `${avatarSize}px`);

    const novelAlign = getMessageAvatarAlign(msg);
    row.dataset.novelAvatarAlign = novelAlign;
    row.dataset.novelTextFlow = getMessageTextFlow(msg);

    const activeContent = getMessageSwipeContent(msg);
    const isExpanded = expandedMessageMenuIndex === index;
    const isEditing = messageInlineEditIndex === index;
    bubble.classList.toggle("actions-expanded", isExpanded || isEditing);

    const menuTrigger = row.querySelector(".message-menu-trigger");
    menuTrigger?.addEventListener("click", (event) => {
      event.stopPropagation();
      expandedMessageMenuIndex = expandedMessageMenuIndex === index ? -1 : index;
      renderMessages();
    });

    row.querySelectorAll(".message-icon-btn").forEach((button) => {
      const action = button.dataset.inlineAction;
      if (!action) return;
      if (action === "edit") {
        button.classList.toggle("hidden", isEditing);
      }
      if (action === "edit-cancel" || action === "edit-save") {
        button.classList.toggle("hidden", !isEditing);
      }
      button.addEventListener("click", (event) => {
        event.stopPropagation();
        if (action === "delete") {
          messageActionState = { index, mode: "delete" };
          startMessageDeleteFlow();
          return;
        }
        if (action === "regenerate") {
          regenerateFromMessageAction(index);
          return;
        }
        if (action === "fork") {
          forkFromMessage(index);
          return;
        }
        if (action === "edit") {
          messageInlineEditIndex = index;
          expandedMessageMenuIndex = index;
          renderMessages();
          return;
        }
        if (action === "edit-cancel") {
          messageInlineEditIndex = -1;
          expandedMessageMenuIndex = -1;
          renderMessages();
          return;
        }
        if (action === "edit-save") {
          const value = row.querySelector(".message-edit-input")?.value ?? "";
          commitInlineMessageEdit(index, value);
        }
      });
    });

    const thinkingChip = row.querySelector(".thinking-chip");
    const thinkingSummary = row.querySelector(".thinking-summary");
    const activeProvider = getActiveProvider();
    if (msg.thinkingSummary && activeProvider?.showThinkingSummaries !== false) {
      thinkingChip?.classList.remove("hidden");
      thinkingSummary.textContent = msg.thinkingSummary;
      thinkingChip?.addEventListener("click", () => thinkingSummary.classList.toggle("hidden"));
    }

    const contentEl = row.querySelector(".message-content");
    const editInput = row.querySelector(".message-edit-input");
    if (isEditing) {
      contentEl.classList.add("hidden");
      editInput.classList.remove("hidden");
      editInput.value = activeContent;
      window.setTimeout(() => {
        editInput.focus();
        editInput.setSelectionRange(editInput.value.length, editInput.value.length);
      }, 0);
    } else {
      contentEl.classList.remove("hidden");
      editInput.classList.add("hidden");
      contentEl.innerHTML = formatMessageContent(activeContent);
    }

    const swipeControls = row.querySelector(".message-swipe-controls");
    if (swipeControls && Array.isArray(msg.swipes) && msg.swipes.length > 1 && !isEditing) {
      const swipeId = clampSwipeIndex(msg.swipe_id, msg.swipes.length);
      msg.swipe_id = swipeId;
      swipeControls.classList.remove("hidden");
      swipeControls.querySelector(".message-swipe-counter").textContent = `${swipeId + 1}/${msg.swipes.length}`;
      swipeControls.querySelector(".message-swipe-left").addEventListener("click", () => setMessageSwipeIndex(index, swipeId - 1));
      swipeControls.querySelector(".message-swipe-right").addEventListener("click", () => setMessageSwipeIndex(index, swipeId + 1));
    }

    const avatarUrl = getMessageAvatar(msg);
    setElementBackgroundImage(avatar, avatarUrl);
    els.messages.appendChild(row);
  }

  if (shouldStickToBottom) {
    els.messages.scrollTop = els.messages.scrollHeight;
  }
  updateChatUiVisibility();
}

function renderCharacters() {
  const query = els.characterSearchInput.value.trim().toLowerCase();
  const characters = state.characters.filter((character) => {
    if (!query) return true;
    return `${character.name} ${character.description || ""} ${character.personality || ""} ${character.scenario || ""}`.toLowerCase().includes(query);
  });

  els.characterList.innerHTML = "";
  characters.forEach((character) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-actions">
        <div class="character-card-avatar"></div>
        <div class="entity-card-header">
          <h3>${escapeHtml(character.name || "Untitled character")}</h3>
          <span class="entity-meta">SillyTavern JSON compatible</span>
        </div>
      </div>
      <div class="entity-meta">${escapeHtml(character.scenario || character.description || "No scenario yet.")}</div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="edit">Edit</button>
        <button class="secondary" data-action="new-chat">New chat</button>
      </div>
    `;
    setElementBackgroundImage(card.querySelector(".character-card-avatar"), getCharacterAvatarUrl(character));
    const actions = card.querySelector(".character-inline-actions");
    card.querySelector('[data-action="toggle-actions"]').addEventListener("click", () => actions.classList.toggle("hidden"));
    card.querySelector('[data-action="edit"]').addEventListener("click", () => editCharacter(character.id));
    card.querySelector('[data-action="new-chat"]').addEventListener("click", () => createCharacterChat(character.id));
    els.characterList.appendChild(card);
  });
}

function createCharacter() {
  showCharacterEditor();
  switchView("characterStudioView");
  setStatus("Character studio");
}

function createCharacterChat(characterId) {
  const character = state.characters.find((item) => item.id === characterId);
  if (!character) {
    showToast("Character not found", "error");
    return;
  }
  state.messages = [];
  const chat = createChatRecord(character.name || "New chat");
  chat.characterId = character.id;
  ensureGreetingMessageForChat(chat);
  state.currentChatId = chat.id;
  state.currentChatTitle = chat.title;
  state.messages = Array.isArray(chat.messages) ? chat.messages.map((message) => ({ ...message })) : [];
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
  renderDrawerState();
  switchView("chatsView");
  showToast("New chat created", "success");
}

async function importCharacters(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const normalized = normalizeImportedCharacter(parsed, file.name);
    showCharacterEditor(normalized);
    switchView("characterStudioView");
    els.characterEditorCard.dataset.importedData = JSON.stringify(normalized.data || parsed);
    setStatus("Character imported into studio");
    showToast("Character imported into studio", "success");
  } catch (error) {
    setStatus("Could not import character JSON");
    showToast(error?.message || "Character import failed", "error");
  } finally {
    event.target.value = "";
  }
}

function editCharacter(characterId) {
  const character = state.characters.find((item) => item.id === characterId);
  if (!character) return;
  showCharacterEditor(character);
  switchView("characterStudioView");
  setStatus(`Editing ${character.name || "character"}`);
}

async function importCharacterFromAvatar(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const imported = await importCharacterCardFromImage(file);
    showCharacterEditor(imported);
    switchView("characterStudioView");
    els.characterEditorCard.dataset.importedData = JSON.stringify(imported.data || {});
    setStatus("Character imported into studio");
    showToast("Character imported into studio", "success");
  } catch (error) {
    setStatus("Could not import avatar image");
    showToast(error?.message || "Avatar import failed", "error");
  } finally {
    event.target.value = "";
  }
}

function showProviderForm(provider = null) {
  els.providerFormCard.classList.remove("hidden");
  els.providerFormCard.dataset.editingId = provider?.id || "";
  els.providerNameInput.value = provider?.name || "";
  els.providerUrlInput.value = provider?.baseUrl || "";
  els.providerApiKeyInput.value = provider?.apiKey || "";
  if (els.providerThinkStartInput) els.providerThinkStartInput.value = provider?.thinkingStartTag || "<thinking>";
  if (els.providerThinkEndInput) els.providerThinkEndInput.value = provider?.thinkingEndTag || "</thinking>";
  if (els.providerShowThinkingInput) els.providerShowThinkingInput.checked = provider?.showThinkingSummaries !== false;
  if (els.providerIncludeThinkingInput) els.providerIncludeThinkingInput.checked = Boolean(provider?.includeThinkingInPrompt);
  els.providerNameInput.focus();
}

function hideProviderForm() {
  els.providerFormCard.classList.add("hidden");
  delete els.providerFormCard.dataset.editingId;
  els.providerNameInput.value = "";
  els.providerUrlInput.value = "";
  els.providerApiKeyInput.value = "";
  if (els.providerThinkStartInput) els.providerThinkStartInput.value = "<thinking>";
  if (els.providerThinkEndInput) els.providerThinkEndInput.value = "</thinking>";
  if (els.providerShowThinkingInput) els.providerShowThinkingInput.checked = true;
  if (els.providerIncludeThinkingInput) els.providerIncludeThinkingInput.checked = false;
  setInlineStatus(els.providerFetchModelsStatus, "");
}

function saveProviderFromForm() {
  const name = els.providerNameInput.value.trim();
  const baseUrl = els.providerUrlInput.value.trim();
  const apiKey = els.providerApiKeyInput.value.trim();
  if (!name || !baseUrl || !apiKey) {
    setStatus("Fill in provider name, base URL, and API key");
    showToast("Provider save failed", "error");
    return;
  }

  const editingId = els.providerFormCard.dataset.editingId;
  const existingIndex = editingId
    ? state.providers.findIndex((provider) => provider.id === editingId)
    : state.providers.findIndex((provider) => provider.name.toLowerCase() === name.toLowerCase());
  const provider = {
    id: existingIndex >= 0 ? state.providers[existingIndex].id : crypto.randomUUID(),
    name,
    baseUrl,
    apiKey,
    thinkingStartTag: els.providerThinkStartInput?.value?.trim() || "<thinking>",
    thinkingEndTag: els.providerThinkEndInput?.value?.trim() || "</thinking>",
    showThinkingSummaries: Boolean(els.providerShowThinkingInput?.checked ?? true),
    includeThinkingInPrompt: Boolean(els.providerIncludeThinkingInput?.checked ?? false),
  };

  if (existingIndex >= 0) {
    state.providers[existingIndex] = provider;
  } else {
    state.providers.unshift(provider);
  }

  state.activeProviderId = provider.id;
  state.endpoint = provider.baseUrl;
  state.apiKey = provider.apiKey;
  persistState();
  renderProviders();
  renderDrawerState();
  syncInputsFromState();
  hideProviderForm();
  setStatus(existingIndex >= 0 ? "Provider updated" : "Provider saved");
  showToast(existingIndex >= 0 ? "Provider updated!" : "Provider created!", "success");
}

async function importProviders(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const providers = Array.isArray(parsed) ? parsed : [parsed];
    providers.forEach((provider, index) => {
      const importedProvider = {
        id: crypto.randomUUID(),
        name: provider.name || "Imported provider",
        baseUrl: provider.baseUrl || provider.endpoint || "",
        apiKey: provider.apiKey || "",
        thinkingStartTag: provider.thinkingStartTag || "<thinking>",
        thinkingEndTag: provider.thinkingEndTag || "</thinking>",
        showThinkingSummaries: provider.showThinkingSummaries !== false,
        includeThinkingInPrompt: Boolean(provider.includeThinkingInPrompt),
      };
      state.providers.unshift(importedProvider);
      if (index === 0 && !state.activeProviderId) {
        state.activeProviderId = importedProvider.id;
      }
    });
    persistState();
    renderProviders();
    renderDrawerState();
    setStatus("Provider imported");
  } catch {
    setStatus("Could not import provider JSON");
  } finally {
    event.target.value = "";
  }
}

function exportProviders() {
  downloadJson("maitavern-providers.json", state.providers.map(({ name, baseUrl, apiKey }) => ({ name, baseUrl, apiKey })));
  setStatus("Providers exported");
}

function exportSingleProvider(providerId) {
  const provider = state.providers.find((item) => item.id === providerId);
  if (!provider) return;
  downloadJson(`${slugify(provider.name || "provider")}.json`, {
    name: provider.name,
    baseUrl: provider.baseUrl,
    apiKey: provider.apiKey,
  });
  setStatus("Provider exported");
  showToast("Provider exported", "success");
}

function renderProviders() {
  const query = els.providerSearchInput.value.trim().toLowerCase();
  const providers = state.providers.filter((provider) => {
    if (!query) return true;
    return `${provider.name} ${provider.baseUrl}`.toLowerCase().includes(query);
  });

  els.providerList.innerHTML = "";
  providers.forEach((provider) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-provider-actions">
        <div>
          <h3>${escapeHtml(provider.name)}</h3>
          <div class="entity-meta">OpenAI-compatible</div>
        </div>
      </div>
      <div class="entity-meta">${escapeHtml(provider.baseUrl)}</div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="use">Use</button>
        <button class="secondary" data-action="edit">Edit</button>
        <button class="secondary" data-action="export">Export</button>
        <button class="secondary" data-action="remove">Delete</button>
      </div>
    `;
    const actions = card.querySelector(".character-inline-actions");
    card.querySelector('[data-action="toggle-provider-actions"]').addEventListener("click", () => actions.classList.toggle("hidden"));
    card.querySelector('[data-action="use"]').addEventListener("click", () => useProvider(provider.id));
    card.querySelector('[data-action="edit"]').addEventListener("click", () => showProviderForm(provider));
    card.querySelector('[data-action="export"]').addEventListener("click", () => exportSingleProvider(provider.id));
    card.querySelector('[data-action="remove"]').addEventListener("click", () => removeProvider(provider.id));
    els.providerList.appendChild(card);
  });

  renderDrawerProviders();
}

function useProvider(providerId) {
  const provider = state.providers.find((item) => item.id === providerId);
  if (!provider) return;
  state.activeProviderId = provider.id;
  state.endpoint = provider.baseUrl;
  state.apiKey = provider.apiKey;
  syncInputsFromState();
  persistState();
  renderDrawerState();
  els.providerPicker?.classList.add("hidden");
  setStatus(`Using provider ${provider.name}`);
}

function removeProvider(providerId) {
  state.providers = state.providers.filter((item) => item.id !== providerId);
  if (state.activeProviderId === providerId) {
    state.activeProviderId = "";
  }
  persistState();
  renderProviders();
  renderDrawerState();
  setStatus("Provider removed");
  showToast("Provider deleted", "success");
}

function renderDrawerProviders() {
  if (!els.drawerProviderList) return;
  els.drawerProviderList.innerHTML = "";
  state.providers.forEach((provider) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `drawer-choice${provider.id === state.activeProviderId ? " active" : ""}`;
    button.innerHTML = `
      <strong>${escapeHtml(provider.name)}</strong>
      <small>${escapeHtml(provider.baseUrl)}</small>
    `;
    button.addEventListener("click", () => {
      useProvider(provider.id);
      showDrawerProviderEditor(provider.id);
    });
    els.drawerProviderList.appendChild(button);
  });
}

function getActiveChatLorebookIds() {
  const chat = getCurrentChat();
  return Array.isArray(chat?.lorebookIds)
    ? [...new Set(chat.lorebookIds.map((item) => String(item || "").trim()).filter(Boolean))]
    : [];
}

function getActiveChatLorebookSettings() {
  const chat = getCurrentChat();
  const source = chat?.lorebookSettings && typeof chat.lorebookSettings === "object"
    ? chat.lorebookSettings
    : defaults.lorebookSettings;

  const attachedBooks = state.lorebooks.filter((book) => lbv2IsLorebookAttachedToChat(book, chat));
  const minScan = attachedBooks.length
    ? Math.min(...attachedBooks.map((book) => Number.parseInt(book.scanDepth ?? source?.scanDepth ?? defaults.lorebookSettings.scanDepth, 10) || defaults.lorebookSettings.scanDepth))
    : Number.parseInt(source?.scanDepth, 10);
  const minBudget = attachedBooks.length
    ? Math.min(...attachedBooks.map((book) => Number.parseInt(book.tokenBudget ?? source?.tokenBudget ?? defaults.lorebookSettings.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget))
    : Number.parseInt(source?.tokenBudget, 10);

  return {
    scanDepth: Math.min(64, Math.max(1, Number.isFinite(minScan) ? minScan : defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.isFinite(minBudget) ? minBudget : defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: source?.recursiveScanning !== false,
  };
}

function saveActiveChatLorebookIds(nextIds = []) {
  const chat = getCurrentChat();
  if (!chat) {
    setInlineStatus(els.drawerLorebookStatus, "Open a chat first");
    return;
  }

  chat.lorebookIds = [...new Set(nextIds.map((item) => String(item || "").trim()).filter(Boolean))];
  syncCurrentChatFromState();
  persistState();
  renderDrawerLorebooks();
}

function renderDrawerLorebooks() {
  if (!els.drawerLorebookList) return;

  const chat = getCurrentChat();
  const isChatOpen = Boolean(chat?.id);
  const query = String(els.drawerLorebookSearchInput?.value || "").trim().toLowerCase();
  const selectedIds = new Set(getActiveChatLorebookIds());
  const visibleLorebooks = state.lorebooks.filter((book) => {
    if (!query) return true;
    return `${book.name || ""} ${book.description || ""}`.toLowerCase().includes(query);
  });

  if (!drawerLorebookSelectedId || !state.lorebooks.some((book) => book.id === drawerLorebookSelectedId)) {
    drawerLorebookSelectedId = [...selectedIds][0] || state.lorebooks[0]?.id || "";
  }

  if (els.drawerLorebookSelect) {
    els.drawerLorebookSelect.innerHTML = "";
    state.lorebooks.forEach((book) => {
      const option = document.createElement("option");
      option.value = book.id;
      option.textContent = book.name || "Untitled lorebook";
      option.selected = book.id === drawerLorebookSelectedId;
      els.drawerLorebookSelect.appendChild(option);
    });
  }

  if (els.drawerLorebookEntries) {
    els.drawerLorebookEntries.innerHTML = "";
    const selectedBook = state.lorebooks.find((book) => book.id === drawerLorebookSelectedId);
    if (!selectedBook) {
      const empty = document.createElement("article");
      empty.className = "entity-card";
      empty.innerHTML = `<div class="entity-meta">No lorebook selected.</div>`;
      els.drawerLorebookEntries.appendChild(empty);
    } else if (!selectedBook.entries?.length) {
      const empty = document.createElement("article");
      empty.className = "entity-card";
      empty.innerHTML = `<div class="entity-meta">No entries in ${escapeHtml(selectedBook.name)}.</div>`;
      els.drawerLorebookEntries.appendChild(empty);
    } else {
      selectedBook.entries.forEach((entry, idx) => {
        const item = document.createElement("article");
        item.className = "entity-card";
        item.innerHTML = `
          <div class="entity-card-header">
            <h3>${escapeHtml(entry?.title || `Entry ${idx + 1}`)}</h3>
            <span class="entity-meta">${Array.isArray(entry?.keys) ? entry.keys.length : 0} key(s)</span>
          </div>
          <div class="entity-meta">${escapeHtml(String(entry?.content || "").slice(0, 180) || "No content")}</div>
        `;
        els.drawerLorebookEntries.appendChild(item);
      });
    }
  }

  els.drawerLorebookList.innerHTML = "";
  if (!isChatOpen) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">Open a chat to attach lorebooks.</div>`;
    els.drawerLorebookList.appendChild(empty);
  } else if (!state.lorebooks.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No lorebooks in library yet.</div>`;
    els.drawerLorebookList.appendChild(empty);
  } else if (!visibleLorebooks.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No lorebooks match your search.</div>`;
    els.drawerLorebookList.appendChild(empty);
  } else {
    visibleLorebooks.forEach((book) => {
      const button = document.createElement("button");
      button.type = "button";
      const active = selectedIds.has(book.id);
      button.className = `drawer-choice${active ? " active" : ""}`;
      button.innerHTML = `
        <strong>${escapeHtml(book.name || "Untitled lorebook")}</strong>
        <small>${book.entries?.length || 0} entries${active ? " • attached" : ""}</small>
      `;
      button.addEventListener("click", () => {
        drawerLorebookSelectedId = book.id;
        if (!getCurrentChat()) {
          setInlineStatus(els.drawerLorebookStatus, "Open a chat first");
          renderDrawerLorebooks();
          return;
        }
        const next = new Set(getActiveChatLorebookIds());
        if (next.has(book.id)) next.delete(book.id);
        else next.add(book.id);
        saveActiveChatLorebookIds([...next]);
        setInlineStatus(els.drawerLorebookStatus, next.has(book.id) ? `Attached ${book.name}` : `Detached ${book.name}`);
      });
      els.drawerLorebookList.appendChild(button);
    });
  }

  const selectedNames = state.lorebooks
    .filter((book) => selectedIds.has(book.id))
    .map((book) => book.name)
    .filter(Boolean);
  if (els.currentLorebookLabel) {
    els.currentLorebookLabel.textContent = `Lorebooks in this chat: ${selectedNames.length ? selectedNames.join(", ") : "none"}`;
  }

  const chatSettings = getActiveChatLorebookSettings();
  if (els.drawerLorebookScanDepthInput) els.drawerLorebookScanDepthInput.value = String(chatSettings.scanDepth);
  if (els.drawerLorebookTokenBudgetInput) els.drawerLorebookTokenBudgetInput.value = String(chatSettings.tokenBudget);
  if (els.drawerLorebookRecursiveInput) els.drawerLorebookRecursiveInput.checked = chatSettings.recursiveScanning;
}

function saveDrawerLorebookSettings() {
  const chat = getCurrentChat();
  if (!chat) {
    setInlineStatus(els.drawerLorebookStatus, "Open a chat first");
    return;
  }

  chat.lorebookSettings = {
    scanDepth: Math.min(64, Math.max(1, Number.parseInt(els.drawerLorebookScanDepthInput?.value || "", 10) || defaults.lorebookSettings.scanDepth)),
    tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(els.drawerLorebookTokenBudgetInput?.value || "", 10) || defaults.lorebookSettings.tokenBudget)),
    recursiveScanning: Boolean(els.drawerLorebookRecursiveInput?.checked ?? true),
  };

  syncCurrentChatFromState();
  persistState();
  renderDrawerLorebooks();
  setInlineStatus(els.drawerLorebookStatus, "Lore settings saved for current chat");
}

function renderPresets() {
  if (!els.presetList) return;
  els.presetList.innerHTML = "";
  state.presets.forEach((preset) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="entity-card-header">
        <h3>${escapeHtml(preset.name)}</h3>
        <span class="entity-meta">${escapeHtml((preset.source || "maitavern").toUpperCase())}</span>
      </div>
      <div class="entity-meta">Temp ${Number(preset.temperature ?? 0.7).toFixed(2)} • Top P ${Number(preset.topP ?? 1).toFixed(2)} • Max ${escapeHtml(preset.maxTokens ?? 300)} • Stream ${preset.stream ? "On" : "Off"}</div>
      <div class="entity-meta">System prompt ${preset.useSystemPrompt !== false ? "On" : "Off"} • Context ${preset.contextTemplate ? "loaded" : "none"} • Instruct ${shouldApplyInstructPreset(preset) ? "loaded" : "none"}</div>
      <div class="entity-meta">Prompt blocks ${Array.isArray(preset.prompts) ? preset.prompts.length : 0} • Prompt order ${Array.isArray(preset.promptOrder) && preset.promptOrder.length ? "loaded" : "none"}</div>
      <div class="entity-meta">${escapeHtml((preset.systemPrompt || "No system prompt").slice(0, 140))}</div>
      <div class="entity-card-actions">
        <button class="secondary" data-action="use">Use</button>
        <button class="secondary" data-action="edit">Edit</button>
        <button class="secondary" data-action="export">Export</button>
        <button class="secondary" data-action="remove">Remove</button>
      </div>
    `;
    card.querySelector('[data-action="use"]').addEventListener("click", () => usePreset(preset.id));
    card.querySelector('[data-action="edit"]').addEventListener("click", () => showPresetForm(preset));
    card.querySelector('[data-action="export"]').addEventListener("click", () => exportPreset(preset.id));
    card.querySelector('[data-action="remove"]').addEventListener("click", () => removePreset(preset.id));
    els.presetList.appendChild(card);
  });

  renderDrawerPresets();
}

function showPresetForm(preset = null) {
  const normalized = preset ? normalizePreset(preset) : normalizePreset({ name: "", systemPrompt: "", source: "maitavern" });
  presetEditorDraft = {
    id: normalized.id,
    prompts: getPromptDraftListFromPreset(normalized),
  };

  els.presetFormCard.classList.remove("hidden");
  els.presetFormCard.dataset.editingId = preset?.id || "";
  els.presetNameInput.value = preset?.name || "";
  els.presetTemperatureInput.value = normalized.temperature;
  els.presetTopPInput.value = normalized.topP;
  els.presetMaxTokensInput.value = normalized.maxTokens;
  els.presetStreamInput.checked = Boolean(normalized.stream);
  els.presetSystemPromptInput.value = normalized.systemPrompt || "";
  els.presetContextTemplateInput.value = normalized.contextTemplate || "";
  els.presetChatStartInput.value = normalized.chatStart || "";
  els.presetExampleSeparatorInput.value = normalized.exampleSeparator || "";
  if (els.presetInstructEnabledInput) els.presetInstructEnabledInput.checked = normalized.instruct.enabled === true;
  els.presetInstructUserInput.value = normalized.instruct.inputSequence || "";
  els.presetInstructAssistantInput.value = normalized.instruct.outputSequence || "";
  els.presetInstructSystemInput.value = normalized.instruct.systemSequence || "";
  els.presetStopSequenceInput.value = normalized.instruct.stopSequence || "";
  if (els.presetUseSystemPromptInput) els.presetUseSystemPromptInput.checked = normalized.useSystemPrompt !== false;
  if (els.presetNamesBehaviorInput) els.presetNamesBehaviorInput.value = normalized.namesBehavior || normalized.instruct?.namesBehavior || "";
  if (els.presetScenarioFormatInput) els.presetScenarioFormatInput.value = normalized.scenarioFormat || "{{scenario}}";
  if (els.presetPersonalityFormatInput) els.presetPersonalityFormatInput.value = normalized.personalityFormat || "{{personality}}";

  renderPresetPromptList();
  hidePresetPromptEditor();
  els.presetNameInput.focus();
}

function hidePresetForm() {
  els.presetFormCard.classList.add("hidden");
  delete els.presetFormCard.dataset.editingId;
  presetEditorDraft = null;
  hidePresetPromptEditor();
  els.presetNameInput.value = "";
  els.presetTemperatureInput.value = "0.7";
  els.presetTopPInput.value = "1";
  els.presetMaxTokensInput.value = "300";
  els.presetStreamInput.checked = Boolean(defaults.presets[0].stream);
  els.presetSystemPromptInput.value = "";
  els.presetContextTemplateInput.value = defaults.presets[0].contextTemplate;
  els.presetChatStartInput.value = defaults.presets[0].chatStart;
  els.presetExampleSeparatorInput.value = defaults.presets[0].exampleSeparator;
  if (els.presetInstructEnabledInput) els.presetInstructEnabledInput.checked = defaults.presets[0].instruct.enabled === true;
  els.presetInstructUserInput.value = defaults.presets[0].instruct.inputSequence;
  els.presetInstructAssistantInput.value = defaults.presets[0].instruct.outputSequence;
  els.presetInstructSystemInput.value = defaults.presets[0].instruct.systemSequence;
  els.presetStopSequenceInput.value = defaults.presets[0].instruct.stopSequence;
  if (els.presetUseSystemPromptInput) els.presetUseSystemPromptInput.checked = true;
  if (els.presetNamesBehaviorInput) els.presetNamesBehaviorInput.value = "";
  if (els.presetScenarioFormatInput) els.presetScenarioFormatInput.value = "{{scenario}}";
  if (els.presetPersonalityFormatInput) els.presetPersonalityFormatInput.value = "{{personality}}";
  if (els.presetPromptList) els.presetPromptList.innerHTML = "";
}

function getPromptDraftListFromPreset(preset) {
  return getOrderedPromptEntries(preset).map((prompt) => ({
    ...prompt,
    enabled: prompt._orderEnabled !== false && prompt.enabled !== false,
  }));
}

function renderPresetPromptList() {
  if (!els.presetPromptList) return;
  const prompts = Array.isArray(presetEditorDraft?.prompts) ? presetEditorDraft.prompts : [];
  els.presetPromptList.innerHTML = "";

  if (!prompts.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No prompt blocks yet. Add one to create toggles/injections like SillyTavern.</div>`;
    els.presetPromptList.appendChild(empty);
    return;
  }

  prompts.forEach((prompt, index) => {
    const card = document.createElement("article");
    card.className = "entity-card";
    card.innerHTML = `
      <div class="character-card-top" data-action="toggle-actions">
        <div class="entity-card-header" style="width:100%; margin-bottom:0;">
          <h3>${escapeHtml(prompt.name || prompt.identifier || `Prompt ${index + 1}`)}</h3>
          <button class="secondary" data-action="enable-toggle" type="button">${prompt.enabled !== false ? "On" : "Off"}</button>
        </div>
      </div>
      <div class="character-inline-actions hidden">
        <button class="secondary" data-action="edit" type="button">Edit</button>
        <button class="danger" data-action="remove" type="button">Delete</button>
      </div>
    `;

    const actions = card.querySelector(".character-inline-actions");

    card.querySelector('[data-action="toggle-actions"]').addEventListener("click", (event) => {
      if (event.target.closest('[data-action="enable-toggle"]')) return;
      actions.classList.toggle("hidden");
    });

    card.querySelector('[data-action="enable-toggle"]').addEventListener("click", (event) => {
      event.stopPropagation();
      prompt.enabled = prompt.enabled === false;
      renderPresetPromptList();
    });

    card.querySelector('[data-action="edit"]').addEventListener("click", () => showPresetPromptEditor(index));

    card.querySelector('[data-action="remove"]').addEventListener("click", () => {
      presetEditorDraft.prompts.splice(index, 1);
      hidePresetPromptEditor();
      renderPresetPromptList();
    });

    els.presetPromptList.appendChild(card);
  });
}

function createPresetPromptDraft() {
  if (!presetEditorDraft) presetEditorDraft = { prompts: [] };
  const nextIndex = (presetEditorDraft.prompts?.length || 0) + 1;
  const prompt = {
    identifier: `custom-${Date.now()}`,
    name: `Custom prompt ${nextIndex}`,
    role: "system",
    content: "",
    marker: false,
    system_prompt: true,
    enabled: true,
    injection_position: 0,
    injection_depth: 4,
    forbid_overrides: false,
  };
  presetEditorDraft.prompts.push(prompt);
  renderPresetPromptList();
  showPresetPromptEditor(presetEditorDraft.prompts.length - 1);
}

function showPresetPromptEditor(index) {
  if (!els.presetPromptEditor || !presetEditorDraft?.prompts?.[index]) return;
  const prompt = presetEditorDraft.prompts[index];
  els.presetPromptModalBackdrop?.classList.remove("hidden");
  els.presetPromptEditor.classList.remove("hidden");
  els.presetPromptEditor.dataset.promptIndex = String(index);
  els.presetPromptIdentifierInput.value = prompt.identifier || "";
  els.presetPromptNameInput.value = prompt.name || "";
  els.presetPromptRoleInput.value = prompt.role || "system";
  els.presetPromptInjectionPositionInput.value = String(Number(prompt.injection_position ?? 0));
  els.presetPromptInjectionDepthInput.value = String(Number(prompt.injection_depth ?? 4));
  els.presetPromptEnabledInput.checked = prompt.enabled !== false;
  els.presetPromptSystemPromptInput.checked = prompt.system_prompt !== false;
  els.presetPromptMarkerInput.checked = Boolean(prompt.marker);
  els.presetPromptForbidOverridesInput.checked = Boolean(prompt.forbid_overrides);
  els.presetPromptContentInput.value = prompt.content || "";
}

function hidePresetPromptEditor() {
  if (!els.presetPromptEditor) return;
  els.presetPromptModalBackdrop?.classList.add("hidden");
  els.presetPromptEditor.classList.add("hidden");
  delete els.presetPromptEditor.dataset.promptIndex;
}

function savePresetPromptDraft() {
  const index = Number(els.presetPromptEditor?.dataset.promptIndex);
  if (!Number.isFinite(index) || !presetEditorDraft?.prompts?.[index]) return;

  const identifier = els.presetPromptIdentifierInput.value.trim();
  const name = els.presetPromptNameInput.value.trim();
  if (!identifier) {
    showToast("Prompt identifier is required", "error");
    return;
  }

  presetEditorDraft.prompts[index] = {
    ...presetEditorDraft.prompts[index],
    identifier,
    name: name || identifier,
    role: els.presetPromptRoleInput.value || "system",
    injection_position: Number.parseInt(els.presetPromptInjectionPositionInput.value, 10) || 0,
    injection_depth: Number.parseInt(els.presetPromptInjectionDepthInput.value, 10) || 0,
    enabled: Boolean(els.presetPromptEnabledInput.checked),
    system_prompt: Boolean(els.presetPromptSystemPromptInput.checked),
    marker: Boolean(els.presetPromptMarkerInput.checked),
    forbid_overrides: Boolean(els.presetPromptForbidOverridesInput.checked),
    content: els.presetPromptContentInput.value,
  };

  renderPresetPromptList();
  hidePresetPromptEditor();
}

function deletePresetPromptDraft() {
  const index = Number(els.presetPromptEditor?.dataset.promptIndex);
  if (!Number.isFinite(index) || !presetEditorDraft?.prompts?.[index]) return;
  presetEditorDraft.prompts.splice(index, 1);
  hidePresetPromptEditor();
  renderPresetPromptList();
}

function savePresetFromForm() {
  const name = els.presetNameInput.value.trim();
  const systemPrompt = els.presetSystemPromptInput.value.trim() || defaults.systemPrompt;
  const temperature = Number.parseFloat(els.presetTemperatureInput.value);
  const topP = Number.parseFloat(els.presetTopPInput.value);
  const maxTokens = Number.parseInt(els.presetMaxTokensInput.value, 10);
  const stream = Boolean(els.presetStreamInput.checked);

  if (!name) {
    setStatus("Add a preset name");
    return;
  }

  if (!Number.isFinite(temperature) || temperature < 0 || temperature > 2) {
    setStatus("Temperature must be between 0 and 2");
    return;
  }

  if (!Number.isFinite(topP) || topP < 0 || topP > 1) {
    setStatus("Top P must be between 0 and 1");
    return;
  }

  if (!Number.isFinite(maxTokens) || maxTokens < 1) {
    setStatus("Max tokens must be greater than 0");
    return;
  }

  const promptDraftList = normalizePresetPromptList(presetEditorDraft?.prompts || []);
  const promptOrder = [{
    character_id: 100001,
    order: promptDraftList.map((prompt) => ({
      identifier: prompt.identifier,
      enabled: prompt.enabled !== false,
    })),
  }];

  const presetData = normalizePreset({
    id: els.presetFormCard.dataset.editingId || crypto.randomUUID(),
    name,
    systemPrompt,
    useSystemPrompt: Boolean(els.presetUseSystemPromptInput?.checked ?? true),
    names_behavior: els.presetNamesBehaviorInput?.value?.trim() || "",
    temperature,
    topP,
    maxTokens,
    stream,
    contextTemplate: els.presetContextTemplateInput.value.trim(),
    chatStart: els.presetChatStartInput.value,
    exampleSeparator: els.presetExampleSeparatorInput.value,
    scenarioFormat: els.presetScenarioFormatInput?.value?.trim() || "{{scenario}}",
    personalityFormat: els.presetPersonalityFormatInput?.value?.trim() || "{{personality}}",
    prompts: promptDraftList,
    prompt_order: promptOrder,
    instruct: {
      enabled: Boolean(els.presetInstructEnabledInput?.checked ?? false),
      inputSequence: els.presetInstructUserInput.value,
      outputSequence: els.presetInstructAssistantInput.value,
      systemSequence: els.presetInstructSystemInput.value,
      stopSequence: els.presetStopSequenceInput.value,
      namesBehavior: els.presetNamesBehaviorInput?.value?.trim() || "",
    },
    source: "maitavern",
  });

  const editingId = els.presetFormCard.dataset.editingId;
  const presetIndex = editingId ? state.presets.findIndex((item) => item.id === editingId) : -1;

  if (presetIndex >= 0) {
    state.presets[presetIndex] = presetData;
    setStatus("Preset updated");
  } else {
    state.presets.unshift(presetData);
    state.activePresetId = presetData.id;
    setStatus("Preset created");
  }

  persistState();
  renderPresets();
  renderDrawerState();
  hidePresetForm();
}

async function importPresets(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;

  try {
    for (const file of files) {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const importedPreset = normalizeImportedPreset(parsed, file.name);
      state.presets.unshift(importedPreset);
    }
    persistState();
    renderPresets();
    renderDrawerState();
    setStatus("Preset imported");
  } catch {
    setStatus("Could not import preset JSON");
  } finally {
    event.target.value = "";
  }
}

function exportPresets() {
  downloadJson("maitavern-presets.json", state.presets);
  setStatus("Presets exported");
}

function exportPreset(presetId) {
  const preset = state.presets.find((item) => item.id === presetId);
  if (!preset) return;
  downloadJson(`${slugify(preset.name || "preset")}.json`, preset);
  setStatus("Preset exported");
}

function normalizeImportedPreset(parsed, fileName = "") {
  if (parsed?.chat_completion_source || parsed?.prompts || parsed?.openai_setting_names) {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported OpenAI preset",
      systemPrompt: parsed.prompts?.find((prompt) => prompt.identifier === "main")?.content
        || parsed.prompts?.find((prompt) => prompt.system_prompt && prompt.content)?.content
        || parsed.system_prompt
        || defaults.systemPrompt,
      use_sysprompt: parsed.use_sysprompt,
      temperature: parsed.temperature,
      top_p: parsed.top_p,
      openai_max_tokens: parsed.openai_max_tokens,
      stream: parsed.stream_openai ?? parsed.stream ?? false,
      scenario_format: parsed.scenario_format || parsed.scenarioFormat,
      personality_format: parsed.personality_format || parsed.personalityFormat,
      prompts: parsed.prompts,
      prompt_order: parsed.prompt_order,
      names_behavior: parsed.names_behavior,
      source: "sillytavern-openai",
      raw: parsed,
    });
  }

  if (parsed?.story_string || parsed?.chat_start || parsed?.example_separator || parsed?.trim_sentences !== undefined) {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported context preset",
      contextTemplate: parsed.story_string,
      chat_start: parsed.chat_start,
      example_separator: parsed.example_separator,
      scenario_format: parsed.scenario_format || parsed.scenarioFormat,
      personality_format: parsed.personality_format || parsed.personalityFormat,
      source: "sillytavern-context",
      raw: parsed,
    });
  }

  if (parsed?.input_sequence || parsed?.output_sequence || parsed?.system_sequence || parsed?.names_behavior !== undefined) {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported instruct preset",
      input_sequence: parsed.input_sequence,
      output_sequence: parsed.output_sequence,
      system_sequence: parsed.system_sequence,
      stop_sequence: parsed.stop_sequence,
      wrap: parsed.wrap,
      macro: parsed.macro,
      names_behavior: parsed.names_behavior,
      input_suffix: parsed.input_suffix,
      output_suffix: parsed.output_suffix,
      system_suffix: parsed.system_suffix,
      first_input_sequence: parsed.first_input_sequence,
      first_output_sequence: parsed.first_output_sequence,
      last_input_sequence: parsed.last_input_sequence,
      last_output_sequence: parsed.last_output_sequence,
      system_same_as_user: parsed.system_same_as_user,
      story_string_prefix: parsed.story_string_prefix,
      story_string_suffix: parsed.story_string_suffix,
      source: "sillytavern-instruct",
      raw: parsed,
    });
  }

  if (typeof parsed?.content === "string") {
    return normalizePreset({
      name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported system prompt",
      systemPrompt: parsed.content,
      source: "sillytavern-sysprompt",
      raw: parsed,
    });
  }

  return normalizePreset({
    ...parsed,
    name: parsed.name || fileName.replace(/\.json$/i, "") || "Imported preset",
    source: parsed.source || "imported",
    raw: parsed,
  });
}

function usePreset(presetId) {
  const preset = state.presets.find((item) => item.id === presetId);
  if (!preset) return;
  state.activePresetId = preset.id;
  persistState();
  renderDrawerState();
  setStatus(`Using preset ${preset.name}`);
}

function removePreset(presetId) {
  if (state.presets.length <= 1) {
    setStatus("Keep at least one preset");
    return;
  }

  state.presets = state.presets.filter((item) => item.id !== presetId);
  if (state.activePresetId === presetId) {
    state.activePresetId = state.presets[0]?.id || defaults.activePresetId;
  }
  persistState();
  renderPresets();
  renderDrawerState();
  setStatus("Preset removed");
}

function renderDrawerPresets() {
  if (!els.drawerPresetList) return;
  els.presetPicker?.classList.remove("hidden");
  els.drawerPresetList.innerHTML = "";
  state.presets.forEach((preset) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `drawer-choice${preset.id === state.activePresetId ? " active" : ""}`;
    button.innerHTML = `
      <strong>${escapeHtml(preset.name)}</strong>
      <small>Temp ${Number(preset.temperature ?? 0.7).toFixed(2)} • Top P ${Number(preset.topP ?? 1).toFixed(2)}</small>
    `;
    button.addEventListener("click", () => {
      usePreset(preset.id);
      showDrawerPresetEditor(state.activePresetId);
      openDrawerPanel("preset-editor");
    });
    els.drawerPresetList.appendChild(button);
  });
}

function openChat(chatId) {
  const chat = state.chats.find((item) => item.id === chatId);
  if (!chat) return;
  ensureGreetingMessageForChat(chat);
  state.currentChatId = chat.id;
  state.currentChatTitle = chat.title || "New chat";
  state.messages = Array.isArray(chat.messages) ? chat.messages.map((message) => ({ ...message })) : [];
  renderMessages();
  renderDrawerState();
  switchView("chatsView");
  updateChatUiVisibility();
  persistState();
  setStatus(`Opened ${state.currentChatTitle}`);
}

function renderRecentChats() {
  els.recentChats.innerHTML = "";
  state.recentChats.forEach((chat) => {
    const character = state.characters.find((item) => item.id === chat.characterId);
    const card = document.createElement("article");
    card.className = "chat-card";
    card.innerHTML = `
      <div class="character-card-top">
        <div class="character-card-avatar"></div>
        <div>
          <h3>${escapeHtml(chat.title)}</h3>
          <div class="entity-meta">${escapeHtml(chat.preview)}</div>
        </div>
      </div>
    `;
    setElementBackgroundImage(card.querySelector(".character-card-avatar"), getCharacterAvatarUrl(character));
    card.addEventListener("click", () => openChat(chat.id));
    els.recentChats.appendChild(card);
  });
}

function updateChatUiVisibility() {
  const isChatView = state.activeView === "chatsView";
  const hasActiveChat = isChatView && Boolean(state.currentChatId);
  const hasChatList = isChatView && state.recentChats.length > 0;
  els.appShell?.classList.toggle("chat-fullscreen", isChatView);
  els.chatHamburgerBtn?.classList.toggle("hidden", !isChatView);
  els.composer.classList.toggle("hidden", !hasActiveChat);
  els.messages.classList.toggle("hidden", !isChatView || !hasActiveChat);
  els.chatEmptyState.classList.toggle("hidden", !isChatView || hasActiveChat || hasChatList);
  els.recentChats.classList.toggle("hidden", !isChatView || hasActiveChat);
  if (!isChatView) toggleChatDrawer(false);
}

function setBusy(isBusy) {
  els.sendBtn.disabled = isBusy;
  els.messageInput.disabled = isBusy;
  els.impersonateSelect.disabled = isBusy;
  els.stopBtn.classList.toggle("hidden", !isBusy);
  if (isBusy) updateComposerOffset();
}

function updateComposerOffset() {
  const composerHeight = els.composer?.offsetHeight || 0;
  document.documentElement.style.setProperty("--composer-offset", `${composerHeight + 24}px`);
}

function buildDrawerPresetDraft(preset) {
  const orderedPrompts = getOrderedPromptEntries(preset).map((prompt) => ({
    ...prompt,
    enabled: prompt._orderEnabled !== false && prompt.enabled !== false,
  }));

  return {
    presetId: preset.id,
    prompts: orderedPrompts,
  };
}

function renderDrawerPresetPromptList() {
  if (!els.drawerPresetPromptList) return;
  const prompts = Array.isArray(drawerPresetEditorDraft?.prompts) ? drawerPresetEditorDraft.prompts : [];
  els.drawerPresetPromptList.innerHTML = "";

  if (!prompts.length) {
    const empty = document.createElement("article");
    empty.className = "entity-card";
    empty.innerHTML = `<div class="entity-meta">No prompt blocks in this preset.</div>`;
    els.drawerPresetPromptList.appendChild(empty);
    return;
  }

  prompts.forEach((prompt, index) => {
    const item = document.createElement("article");
    item.className = "entity-card drawer-prompt-item";
    item.draggable = true;
    item.dataset.promptIndex = String(index);
    item.innerHTML = `
      <div class="drawer-prompt-top">
        <div>
          <h3>${escapeHtml(prompt.name || prompt.identifier || `Prompt ${index + 1}`)}</h3>
          <div class="entity-meta">${escapeHtml(prompt.identifier || "")}</div>
        </div>
        <button class="secondary" data-action="toggle" type="button">${prompt.enabled !== false ? "On" : "Off"}</button>
      </div>
      <div class="entity-meta">Role ${escapeHtml(prompt.role || "system")} • Pos ${escapeHtml(String(prompt.injection_position ?? 0))} • Depth ${escapeHtml(String(prompt.injection_depth ?? 0))}</div>
      <div class="entity-meta">⇅ Drag to reorder</div>
    `;

    item.querySelector('[data-action="toggle"]')?.addEventListener("click", (event) => {
      event.stopPropagation();
      prompt.enabled = prompt.enabled === false;
      renderDrawerPresetPromptList();
    });

    item.addEventListener("dragstart", (event) => {
      drawerPresetPromptDragIndex = index;
      item.classList.add("is-dragging");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", String(index));
      }
    });

    item.addEventListener("dragover", (event) => {
      event.preventDefault();
      item.classList.add("drag-over");
    });

    item.addEventListener("dragleave", () => {
      item.classList.remove("drag-over");
    });

    item.addEventListener("drop", (event) => {
      event.preventDefault();
      item.classList.remove("drag-over");
      const fromIndex = drawerPresetPromptDragIndex;
      const toIndex = index;
      if (!Number.isInteger(fromIndex) || fromIndex < 0 || fromIndex >= prompts.length || fromIndex === toIndex) {
        return;
      }
      const [moved] = prompts.splice(fromIndex, 1);
      prompts.splice(toIndex, 0, moved);
      drawerPresetPromptDragIndex = -1;
      renderDrawerPresetPromptList();
    });

    item.addEventListener("dragend", () => {
      drawerPresetPromptDragIndex = -1;
      item.classList.remove("is-dragging", "drag-over");
      els.drawerPresetPromptList
        ?.querySelectorAll(".drawer-prompt-item.drag-over")
        ?.forEach((node) => node.classList.remove("drag-over"));
    });

    els.drawerPresetPromptList.appendChild(item);
  });
}

function showDrawerPresetEditor(presetId) {
  const preset = state.presets.find((item) => item.id === presetId) || getActivePreset();
  if (!preset || !els.drawerPresetEditor) return;
  els.drawerPresetEditor.classList.remove("hidden");
  els.drawerPresetEditor.dataset.presetId = preset.id;
  drawerPresetEditorDraft = buildDrawerPresetDraft(preset);
  if (els.drawerPresetNameInput) els.drawerPresetNameInput.value = preset.name || "";
  if (els.drawerPresetTempInput) els.drawerPresetTempInput.value = String(preset.temperature ?? 0.7);
  if (els.drawerPresetTopPInput) els.drawerPresetTopPInput.value = String(preset.topP ?? 1);
  if (els.drawerPresetSystemInput) els.drawerPresetSystemInput.value = preset.systemPrompt || "";
  renderDrawerPresetPromptList();
}

function hideDrawerPresetEditor() {
  drawerPresetPromptDragIndex = -1;
  drawerPresetEditorDraft = null;
  if (els.drawerPresetEditor) {
    els.drawerPresetEditor.classList.add("hidden");
    delete els.drawerPresetEditor.dataset.presetId;
  }

  if (els.chatDrawer?.classList.contains("open") && els.chatDrawer?.dataset?.panel === "preset-editor") {
    closeDrawerPanel();
  }
}

function saveDrawerPresetEdits() {
  const presetId = drawerPresetEditorDraft?.presetId || els.drawerPresetEditor?.dataset?.presetId || state.activePresetId;
  const presetIndex = state.presets.findIndex((item) => item.id === presetId);
  if (presetIndex < 0) return;

  const current = state.presets[presetIndex];
  const nextTemp = Number.parseFloat(els.drawerPresetTempInput?.value || String(current.temperature ?? 0.7));
  const nextTopP = Number.parseFloat(els.drawerPresetTopPInput?.value || String(current.topP ?? 1));

  if (!Number.isFinite(nextTemp) || nextTemp < 0 || nextTemp > 2) {
    showToast("Temperature must be between 0 and 2", "error");
    return;
  }
  if (!Number.isFinite(nextTopP) || nextTopP < 0 || nextTopP > 1) {
    showToast("Top P must be between 0 and 1", "error");
    return;
  }

  const promptDraftList = normalizePresetPromptList(drawerPresetEditorDraft?.prompts || current.prompts || []);
  const promptOrder = [{
    character_id: 100001,
    order: promptDraftList.map((prompt) => ({
      identifier: prompt.identifier,
      enabled: prompt.enabled !== false,
    })),
  }];

  const updatedPreset = normalizePreset({
    ...current,
    name: current.name,
    temperature: nextTemp,
    topP: nextTopP,
    systemPrompt: els.drawerPresetSystemInput?.value?.trim() || current.systemPrompt,
    prompts: promptDraftList,
    promptOrder,
  });

  state.presets[presetIndex] = updatedPreset;
  if (state.activePresetId === updatedPreset.id) {
    showDrawerPresetEditor(updatedPreset.id);
  }

  persistState();
  renderPresets();
  renderDrawerState();
  setStatus("Preset updated");
  showToast("Preset updated", "success");
}

function showDrawerProviderEditor(providerId) {
  const provider = state.providers.find((item) => item.id === providerId || item.id === state.activeProviderId);
  if (!provider || !els.drawerProviderEditor) return;
  els.drawerProviderEditor.classList.remove("hidden");
  els.drawerProviderEditor.dataset.providerId = provider.id;
  if (els.drawerProviderNameInput) els.drawerProviderNameInput.value = provider.name || "";
  if (els.drawerProviderUrlInput) els.drawerProviderUrlInput.value = provider.baseUrl || "";
  if (els.drawerProviderApiKeyInput) els.drawerProviderApiKeyInput.value = provider.apiKey || "";
  if (els.drawerProviderThinkStartInput) els.drawerProviderThinkStartInput.value = provider.thinkingStartTag || "<thinking>";
  if (els.drawerProviderThinkEndInput) els.drawerProviderThinkEndInput.value = provider.thinkingEndTag || "</thinking>";
  if (els.drawerProviderShowThinkingInput) els.drawerProviderShowThinkingInput.checked = provider.showThinkingSummaries !== false;
  if (els.drawerProviderIncludeThinkingInput) els.drawerProviderIncludeThinkingInput.checked = Boolean(provider.includeThinkingInPrompt);
}

function hideDrawerProviderEditor() {
  if (!els.drawerProviderEditor) return;
  els.drawerProviderEditor.classList.add("hidden");
  delete els.drawerProviderEditor.dataset.providerId;
}

function saveDrawerProviderEdits() {
  const providerId = els.drawerProviderEditor?.dataset?.providerId || state.activeProviderId;
  const providerIndex = state.providers.findIndex((item) => item.id === providerId);
  if (providerIndex < 0) return;
  const current = state.providers[providerIndex];
  const updated = {
    ...current,
    name: els.drawerProviderNameInput?.value?.trim() || current.name,
    baseUrl: els.drawerProviderUrlInput?.value?.trim() || current.baseUrl,
    apiKey: els.drawerProviderApiKeyInput?.value?.trim() || current.apiKey,
    thinkingStartTag: els.drawerProviderThinkStartInput?.value?.trim() || current.thinkingStartTag || "<thinking>",
    thinkingEndTag: els.drawerProviderThinkEndInput?.value?.trim() || current.thinkingEndTag || "</thinking>",
    showThinkingSummaries: Boolean(els.drawerProviderShowThinkingInput?.checked ?? true),
    includeThinkingInPrompt: Boolean(els.drawerProviderIncludeThinkingInput?.checked ?? false),
  };
  state.providers[providerIndex] = updated;
  if (updated.id === state.activeProviderId) {
    state.endpoint = updated.baseUrl;
    state.apiKey = updated.apiKey;
    syncInputsFromState();
  }
  persistState();
  renderProviders();
  renderDrawerState();
  setStatus("Provider updated");
}

function openMessageActionPopup(index) {
  messageActionState = { index, mode: "delete" };
  if (!els.messageActionPopup) return;
  els.messageActionPopup.classList.remove("hidden");
}

function closeMessageActionPopup() {
  if (!els.messageActionPopup) return;
  els.messageActionPopup.classList.add("hidden");
  messageActionState = { index: -1, mode: "" };
}

function commitInlineMessageEdit(index, value) {
  const msg = state.messages[index];
  if (!msg) return;
  const nextText = String(value ?? "");
  msg.content = nextText;
  if (Array.isArray(msg.swipes) && msg.swipes.length) {
    const swipeId = clampSwipeIndex(Number(msg.swipe_id ?? 0), msg.swipes.length);
    msg.swipes[swipeId] = nextText;
    msg.swipe_id = swipeId;
  }
  messageInlineEditIndex = -1;
  expandedMessageMenuIndex = -1;
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
}

function forkFromMessage(index) {
  if (index < 0 || index >= state.messages.length) return;
  const chat = createChatRecord(`${state.currentChatTitle || "Chat"} (fork)`);
  chat.characterId = getCurrentChat()?.characterId || "";
  chat.messages = state.messages.slice(0, index + 1).map((message) => ({ ...message }));
  state.currentChatId = chat.id;
  state.currentChatTitle = chat.title;
  state.messages = chat.messages.map((message) => ({ ...message }));
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
  renderDrawerState();
  showToast("Forked chat from this message", "success");
}

function startMessageDeleteFlow() {
  const index = messageActionState.index;
  if (index < 0 || index >= state.messages.length) return;
  const below = Math.max(0, state.messages.length - index - 1);
  if (els.messageDeletePromptText) {
    els.messageDeletePromptText.textContent = below
      ? `Delete this message, or delete ${below} message(s) below?`
      : "Delete this message";
  }
  if (els.deleteFromHereBtn) {
    els.deleteFromHereBtn.textContent = below > 0 ? `Delete (${below}) below` : "Delete below";
    els.deleteFromHereBtn.disabled = below < 1;
  }
  expandedMessageMenuIndex = -1;
  openMessageActionPopup(index);
}

function commitMessageDeleteFlow(deleteBelow = false) {
  const index = messageActionState.index;
  if (index < 0 || index >= state.messages.length) return;
  if (deleteBelow) {
    state.messages = state.messages.slice(0, index);
  } else {
    state.messages.splice(index, 1);
  }
  messageInlineEditIndex = -1;
  expandedMessageMenuIndex = -1;
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
  closeMessageActionPopup();
}

async function regenerateFromMessageAction(index = messageActionState.index) {
  if (index < 0 || index >= state.messages.length) return;
  const msg = state.messages[index];
  if (!msg || msg.role !== "assistant") {
    showToast("Regenerate is available on assistant messages", "error");
    return;
  }

  state.messages = state.messages.slice(0, index);
  const assistantMessage = { role: "assistant", content: "", pending: true };
  state.messages.push(assistantMessage);
  const requestMessages = state.messages.slice(0, -1);

  expandedMessageMenuIndex = -1;
  messageInlineEditIndex = -1;
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  closeMessageActionPopup();
  await generateAssistantReply({ requestMessages, assistantMessage });
}

function setStatus(text) {
  els.statusText.textContent = text;
}

function showToast(message, type = "success") {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  els.toastContainer.appendChild(toast);
  window.setTimeout(() => toast.remove(), 2600);
}

function persistState() {
  // Source of truth: local device file named `user`.
  persistUserFile();
}

window.addEventListener("beforeunload", () => {
  try {
    if (pendingUserFilePayload && hasFileSystemAccessApi()) {
      // best-effort flush request
      void flushPendingUserFileWrite();
    }
  } catch {
    // no-op
  }
});

function showCharacterEditor(character = null) {
  const normalized = character ? normalizeCharacterRecord(character) : {
    id: "",
    name: "",
    scenario: "",
    personality: "",
    greeting: "",
    alternateGreetings: [],
    avatarDataUrl: "",
  };
  els.characterEditorCard.dataset.editingId = character?.id || "";
  els.characterEditorCard.dataset.avatarDataUrl = normalized?.avatarDataUrl || "";
  els.characterNameInput.value = normalized?.name || "";
  els.characterScenarioInput.value = normalized?.scenario || "";
  els.characterPersonalityInput.value = normalized?.personality || "";
  els.characterGreetingInput.value = normalized?.greeting || "";
  if (els.characterAlternateGreetingsInput) {
    els.characterAlternateGreetingsInput.value = formatAlternateGreetingsForEditor(normalized?.alternateGreetings || []);
  }
  renderCharacterAvatarPreview(normalized?.avatarDataUrl || "");
  els.characterNameInput.focus();
}

function hideCharacterEditor() {
  delete els.characterEditorCard.dataset.editingId;
  delete els.characterEditorCard.dataset.avatarDataUrl;
  delete els.characterEditorCard.dataset.importedData;
  els.characterNameInput.value = "";
  els.characterScenarioInput.value = "";
  els.characterPersonalityInput.value = "";
  els.characterGreetingInput.value = "";
  if (els.characterAlternateGreetingsInput) {
    els.characterAlternateGreetingsInput.value = "";
  }
  els.characterAvatarInput.value = "";
  renderCharacterAvatarPreview("");
  switchView("charactersView");
}

async function onCharacterAvatarSelected(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const avatarDataUrl = await readFileAsDataUrl(file);
    els.characterEditorCard.dataset.avatarDataUrl = avatarDataUrl;
    renderCharacterAvatarPreview(avatarDataUrl);
  } catch {
    setStatus("Could not load avatar image");
    showToast("Avatar load failed", "error");
  }
}

function saveCharacterFromEditor() {
  const name = els.characterNameInput.value.trim();
  const scenario = els.characterScenarioInput.value.trim();
  const personality = els.characterPersonalityInput.value.trim();
  const greeting = els.characterGreetingInput.value.trim();
  const alternateGreetings = parseAlternateGreetings(els.characterAlternateGreetingsInput?.value || "", { fromEditor: true });
  const avatarDataUrl = els.characterEditorCard.dataset.avatarDataUrl || "";
  if (!name) {
    setStatus("Add a character name");
    showToast("Character save failed", "error");
    return;
  }

  const editingId = els.characterEditorCard.dataset.editingId;
  const character = editingId ? state.characters.find((item) => item.id === editingId) : null;
  const characterPayload = {
    name,
    scenario,
    personality,
    first_mes: greeting,
    alternate_greetings: alternateGreetings,
    description: personality || scenario || "",
    avatarDataUrl,
    spec: "chara_card_v2",
  };

  if (character) {
    character.name = name;
    character.scenario = scenario;
    character.personality = personality;
    character.greeting = greeting;
    character.alternateGreetings = alternateGreetings;
    character.description = personality || scenario || "";
    character.avatarDataUrl = avatarDataUrl;
    character.data = {
      ...(character.data && typeof character.data === "object" ? character.data : {}),
      ...characterPayload,
    };
    setStatus("Character updated");
    showToast("Character saved!", "success");
  } else {
    state.characters.unshift(normalizeCharacterRecord({
      id: crypto.randomUUID(),
      name,
      scenario,
      personality,
      greeting,
      alternateGreetings,
      description: personality || scenario || "",
      avatarDataUrl,
      data: characterPayload,
    }));
    setStatus("Character created");
    showToast("Character saved!", "success");
  }

  persistState();
  renderCharacters();
  renderImpersonationOptions();
  renderDrawerState();
  hideCharacterEditor();
  showToast("Character saved!", "success");
}

function renderCharacterAvatarPreview(avatarDataUrl) {
  els.characterAvatarPreview.classList.toggle("hidden", !avatarDataUrl);
  setElementBackgroundImage(els.characterAvatarPreview, avatarDataUrl);
}

function syncCustomizationInputs() {
  if (els.chatStyleSelect) els.chatStyleSelect.value = state.customization.chatStyle || "bubble";
  if (els.avatarSizeInput) els.avatarSizeInput.value = state.customization.avatarSize;
  if (els.defaultTextColorInput) els.defaultTextColorInput.value = state.customization.defaultTextColor;
  if (els.wrapRuleSearchInput) els.wrapRuleSearchInput.value = wrapRuleSearchQuery;

  if (els.bubbleAssistantColorInput) els.bubbleAssistantColorInput.value = state.customization.bubble.assistantBubbleColor;
  if (els.bubbleUserColorInput) els.bubbleUserColorInput.value = state.customization.bubble.userBubbleColor;
  if (els.bubbleShowAssistantNameInput) els.bubbleShowAssistantNameInput.checked = state.customization.bubble.showAssistantName !== false;
  if (els.bubbleShowUserNameInput) els.bubbleShowUserNameInput.checked = state.customization.bubble.showUserName !== false;
  if (els.bubbleShowAssistantAvatarInput) els.bubbleShowAssistantAvatarInput.checked = state.customization.bubble.showAssistantAvatar !== false;
  if (els.bubbleShowUserAvatarInput) els.bubbleShowUserAvatarInput.checked = state.customization.bubble.showUserAvatar !== false;
  if (els.bubbleAssistantAvatarSizeInput) els.bubbleAssistantAvatarSizeInput.value = String(state.customization.bubble.assistantAvatarSize || 40);
  if (els.bubbleUserAvatarSizeInput) els.bubbleUserAvatarSizeInput.value = String(state.customization.bubble.userAvatarSize || 40);
  if (els.bubbleAssistantAvatarShapeInput) els.bubbleAssistantAvatarShapeInput.value = state.customization.bubble.assistantAvatarShape || "cutout";
  if (els.bubbleUserAvatarShapeInput) els.bubbleUserAvatarShapeInput.value = state.customization.bubble.userAvatarShape || "cutout";

  if (els.novelAssistantColorInput) els.novelAssistantColorInput.value = state.customization.novel.assistantBackgroundColor;
  if (els.novelUserColorInput) els.novelUserColorInput.value = state.customization.novel.userBackgroundColor;
  if (els.novelChatBackgroundColorInput) els.novelChatBackgroundColorInput.value = state.customization.novel.chatBackgroundColor;
  if (els.novelAssistantBlurInput) els.novelAssistantBlurInput.value = String(state.customization.novel.assistantBackgroundBlur || 0);
  if (els.novelUserBlurInput) els.novelUserBlurInput.value = String(state.customization.novel.userBackgroundBlur || 0);
  if (els.novelChatBackgroundBlurInput) els.novelChatBackgroundBlurInput.value = String(state.customization.novel.chatBackgroundBlur || 0);
  if (els.novelShowAssistantNameInput) els.novelShowAssistantNameInput.checked = state.customization.novel.showAssistantName !== false;
  if (els.novelShowUserNameInput) els.novelShowUserNameInput.checked = state.customization.novel.showUserName !== false;
  if (els.novelShowAssistantAvatarInput) els.novelShowAssistantAvatarInput.checked = state.customization.novel.showAssistantAvatar !== false;
  if (els.novelShowUserAvatarInput) els.novelShowUserAvatarInput.checked = state.customization.novel.showUserAvatar !== false;
  if (els.novelAssistantAvatarSizeInput) els.novelAssistantAvatarSizeInput.value = String(state.customization.novel.assistantAvatarSize || 64);
  if (els.novelUserAvatarSizeInput) els.novelUserAvatarSizeInput.value = String(state.customization.novel.userAvatarSize || 64);
  if (els.novelAssistantAvatarShapeInput) els.novelAssistantAvatarShapeInput.value = state.customization.novel.assistantAvatarShape || "cutout";
  if (els.novelUserAvatarShapeInput) els.novelUserAvatarShapeInput.value = state.customization.novel.userAvatarShape || "cutout";
  if (els.novelAssistantAvatarAlignInput) els.novelAssistantAvatarAlignInput.value = state.customization.novel.assistantAvatarAlign || "left";
  if (els.novelUserAvatarAlignInput) els.novelUserAvatarAlignInput.value = state.customization.novel.userAvatarAlign || "left";
  if (els.novelAssistantTextFlowInput) els.novelAssistantTextFlowInput.value = state.customization.novel.assistantTextFlow || "wrap";
  if (els.novelUserTextFlowInput) els.novelUserTextFlowInput.value = state.customization.novel.userTextFlow || "wrap";

  if (els.novelAssistantImageMeta) {
    els.novelAssistantImageMeta.textContent = state.customization.novel.assistantBackgroundImage
      ? "Assistant image: selected"
      : "Assistant image: none";
  }
  if (els.novelUserImageMeta) {
    els.novelUserImageMeta.textContent = state.customization.novel.userBackgroundImage
      ? "User image: selected"
      : "User image: none";
  }
  if (els.novelChatImageMeta) {
    els.novelChatImageMeta.textContent = state.customization.novel.chatBackgroundImage
      ? "Chat image: selected"
      : "Chat image: none";
  }

  if (els.googleFontSearchInput && state.customization.appFontSource === "google") {
    els.googleFontSearchInput.value = state.customization.appFontGoogleName || state.customization.appFontFamily || "";
  }
  if (els.uploadedFontMeta) {
    if (state.customization.appFontSource === "file" && state.customization.appFontFileName) {
      els.uploadedFontMeta.textContent = `Uploaded: ${state.customization.appFontFileName}`;
    } else if (state.customization.appFontSource === "google" && state.customization.appFontFamily) {
      els.uploadedFontMeta.textContent = `Google Font: ${state.customization.appFontFamily}`;
    } else {
      els.uploadedFontMeta.textContent = "Using system default font.";
    }
  }
  syncChatStyleSectionVisibility();
  updateCurrentFontLabel();
}

function getCurrentFontDisplayName() {
  if (state.customization.appFontSource === "file") {
    return state.customization.appFontFamily || state.customization.appFontFileName || "Custom file font";
  }
  if (state.customization.appFontSource === "google") {
    return state.customization.appFontFamily || state.customization.appFontGoogleName || "Google Font";
  }
  return "System default (Inter stack)";
}

function updateCurrentFontLabel() {
  if (!els.currentFontName) return;
  els.currentFontName.textContent = getCurrentFontDisplayName();
}

function renderWrapRuleList() {
  if (!els.wrapRuleList) return;
  const query = wrapRuleSearchQuery.trim().toLowerCase();
  const rules = state.customization.wrapRules.filter((rule) => {
    if (!query) return true;
    const haystack = `${rule.title || ""} ${rule.start || ""} ${rule.end || ""}`.toLowerCase();
    return haystack.includes(query);
  });

  els.wrapRuleList.innerHTML = "";
  if (!rules.length) {
    const empty = document.createElement("p");
    empty.className = "entity-meta wrap-list-empty-state";
    empty.textContent = query ? "No wrap rules match your search." : "No wrap rules yet.";
    els.wrapRuleList.appendChild(empty);
    return;
  }

  rules.forEach((rule) => {
    const item = document.createElement("article");
    item.className = "entity-card";
    item.innerHTML = `
      <div class="entity-card-header">
        <h3>${escapeHtml(rule.title || `${rule.start}${rule.end}`)}</h3>
        <span class="entity-meta">${escapeHtml(rule.start)} ... ${escapeHtml(rule.end)}</span>
      </div>
      <div class="entity-card-actions">
        <span class="wrap-color-dot" aria-hidden="true" style="background: ${escapeAttribute(rule.color || "#c084fc")}"></span>
        <button class="secondary" type="button" data-action="edit-wrap">Edit</button>
      </div>
    `;
    item.querySelector('[data-action="edit-wrap"]')?.addEventListener("click", () => showWrapRuleEditor(rule.id));
    els.wrapRuleList.appendChild(item);
  });
}

function onWrapRuleSearchInput() {
  wrapRuleSearchQuery = els.wrapRuleSearchInput?.value || "";
  renderWrapRuleList();
}

function createWrapRule() {
  const rule = {
    id: crypto.randomUUID(),
    title: "New wrap rule",
    start: "(",
    end: ")",
    color: "#c084fc",
    enabled: true,
  };
  wrapRuleSearchQuery = "";
  if (els.wrapRuleSearchInput) els.wrapRuleSearchInput.value = "";
  state.customization.wrapRules.unshift(rule);
  persistState();
  renderWrapRuleList();
  showWrapRuleEditor(rule.id);
}

function showWrapRuleEditor(ruleId) {
  const rule = state.customization.wrapRules.find((item) => item.id === ruleId);
  if (!rule || !els.wrapRuleEditor) return;
  if (els.textWrappingSection?.classList.contains("hidden")) {
    els.textWrappingSection.classList.remove("hidden");
  }
  els.wrapRuleEditor.classList.remove("hidden");
  els.wrapRuleEditor.dataset.ruleId = rule.id;
  els.wrapRuleTitleInput.value = rule.title || "";
  els.wrapRuleStartInput.value = rule.start || "";
  els.wrapRuleEndInput.value = rule.end || "";
  els.wrapRuleColorInput.value = isValidHexColor(rule.color) ? rule.color : "#c084fc";
  els.wrapRuleHexInput.value = rule.color || "#c084fc";
  els.wrapRuleTitleInput?.focus();
}

function hideWrapRuleEditor() {
  if (!els.wrapRuleEditor) return;
  els.wrapRuleEditor.classList.add("hidden");
  delete els.wrapRuleEditor.dataset.ruleId;
  if (els.wrapRuleTitleInput) els.wrapRuleTitleInput.value = "";
  if (els.wrapRuleStartInput) els.wrapRuleStartInput.value = "";
  if (els.wrapRuleEndInput) els.wrapRuleEndInput.value = "";
  if (els.wrapRuleColorInput) els.wrapRuleColorInput.value = "#c084fc";
  if (els.wrapRuleHexInput) els.wrapRuleHexInput.value = "";
}

function saveWrapRuleFromEditor() {
  const rule = state.customization.wrapRules.find((item) => item.id === els.wrapRuleEditor.dataset.ruleId);
  if (!rule) return;
  const start = els.wrapRuleStartInput.value;
  const end = els.wrapRuleEndInput.value;
  if (!start || !end) {
    showToast("Wrap start and end are required", "error");
    return;
  }
  rule.start = start;
  rule.end = end;
  rule.color = isValidHexColor(els.wrapRuleHexInput.value.trim()) ? els.wrapRuleHexInput.value.trim() : els.wrapRuleColorInput.value;
  rule.title = els.wrapRuleTitleInput.value.trim() || `${rule.start}...${rule.end}`;
  persistState();
  renderWrapRuleList();
  applyCustomization();
  renderMessages();
  hideWrapRuleEditor();
  showToast("Wrap updated!", "success");
}

function deleteWrapRuleFromEditor() {
  const ruleId = els.wrapRuleEditor?.dataset.ruleId;
  if (!ruleId) return;
  state.customization.wrapRules = state.customization.wrapRules.filter((item) => item.id !== ruleId);
  persistState();
  renderWrapRuleList();
  applyCustomization();
  renderMessages();
  hideWrapRuleEditor();
  showToast("Wrap deleted", "success");
}

function toggleSection(element) {
  if (!element) return;
  element.classList.toggle("hidden");
  if (element.classList.contains("hidden")) return;
  if (
    element === els.chatCustomizationSection
    || element === els.textWrappingSection
    || element === els.textFontSection
  ) {
    syncCustomizationInputs();
    renderWrapRuleList();
  }
  if (element === els.textFontSection) {
    toggleFontManager(true);
  }
}

function applyCustomization() {
  const custom = state.customization;
  const bubble = custom.bubble || defaults.customization.bubble;
  const novel = custom.novel || defaults.customization.novel;
  const style = custom.chatStyle === "novel" ? "novel" : "bubble";

  document.documentElement.dataset.chatStyle = style;
  document.documentElement.style.setProperty("--chat-avatar-size", `${custom.avatarSize}px`);
  document.documentElement.style.setProperty("--chat-default-text-color", custom.defaultTextColor);

  document.documentElement.style.setProperty("--chat-bubble-assistant-bg", bubble.assistantBubbleColor || defaults.customization.bubble.assistantBubbleColor);
  document.documentElement.style.setProperty("--chat-bubble-user-bg", bubble.userBubbleColor || defaults.customization.bubble.userBubbleColor);
  document.documentElement.style.setProperty("--chat-bubble-assistant-avatar-size", `${bubble.assistantAvatarSize || custom.avatarSize}px`);
  document.documentElement.style.setProperty("--chat-bubble-user-avatar-size", `${bubble.userAvatarSize || custom.avatarSize}px`);

  document.documentElement.style.setProperty("--chat-novel-assistant-bg", novel.assistantBackgroundColor || defaults.customization.novel.assistantBackgroundColor);
  document.documentElement.style.setProperty("--chat-novel-user-bg", novel.userBackgroundColor || defaults.customization.novel.userBackgroundColor);
  document.documentElement.style.setProperty("--chat-novel-bg", novel.chatBackgroundColor || defaults.customization.novel.chatBackgroundColor);
  document.documentElement.style.setProperty("--chat-novel-assistant-avatar-size", `${novel.assistantAvatarSize || 64}px`);
  document.documentElement.style.setProperty("--chat-novel-user-avatar-size", `${novel.userAvatarSize || 64}px`);
  const assistantBlur = Math.max(0, Number(novel.assistantBackgroundBlur || 0));
  const userBlur = Math.max(0, Number(novel.userBackgroundBlur || 0));
  const chatBlur = Math.max(0, Number(novel.chatBackgroundBlur || 0));

  document.documentElement.style.setProperty("--chat-novel-assistant-filter", assistantBlur > 0 ? `blur(${assistantBlur}px)` : "none");
  document.documentElement.style.setProperty("--chat-novel-user-filter", userBlur > 0 ? `blur(${userBlur}px)` : "none");
  document.documentElement.style.setProperty("--chat-novel-bg-filter", chatBlur > 0 ? `blur(${chatBlur}px)` : "none");

  document.documentElement.style.setProperty("--chat-novel-assistant-avatar-align", novel.assistantAvatarAlign || "left");
  document.documentElement.style.setProperty("--chat-novel-user-avatar-align", novel.userAvatarAlign || "left");

  const assistantWidth = novel.assistantTextFlow === "below"
    ? "100%"
    : "calc(100% - var(--chat-novel-assistant-avatar-size, 64px) - 10px)";
  const userWidth = novel.userTextFlow === "below"
    ? "100%"
    : "calc(100% - var(--chat-novel-user-avatar-size, 64px) - 10px)";
  document.documentElement.style.setProperty("--chat-novel-assistant-content-width", assistantWidth);
  document.documentElement.style.setProperty("--chat-novel-user-content-width", userWidth);

  const assistantBackgroundImage = novel.assistantBackgroundImage
    ? `url(${JSON.stringify(novel.assistantBackgroundImage)})`
    : novel.assistantBackgroundColor;
  const userBackgroundImage = novel.userBackgroundImage
    ? `url(${JSON.stringify(novel.userBackgroundImage)})`
    : novel.userBackgroundColor;
  const chatBackgroundImage = novel.chatBackgroundImage
    ? `url(${JSON.stringify(novel.chatBackgroundImage)})`
    : novel.chatBackgroundColor;

  document.documentElement.style.setProperty("--chat-novel-assistant-bg-image", assistantBackgroundImage);
  document.documentElement.style.setProperty("--chat-novel-user-bg-image", userBackgroundImage);
  document.documentElement.style.setProperty("--chat-novel-bg-image", chatBackgroundImage);

  applyWrapRuleCssVariables();
  applyAppFont();
  updateCurrentFontLabel();
}

function collectCustomizationFromInputs() {
  if (!els.avatarSizeInput || !els.defaultTextColorInput) {
    return false;
  }

  const bubble = state.customization.bubble || {};
  const novel = state.customization.novel || {};

  state.customization.chatStyle = (els.chatStyleSelect?.value === "novel") ? "novel" : "bubble";
  state.customization.avatarSize = Math.min(120, Math.max(24, Number.parseInt(els.avatarSizeInput.value, 10) || defaults.customization.avatarSize));
  state.customization.defaultTextColor = els.defaultTextColorInput.value || defaults.customization.defaultTextColor;

  bubble.assistantBubbleColor = isValidHexColor(els.bubbleAssistantColorInput?.value) ? els.bubbleAssistantColorInput.value : defaults.customization.bubble.assistantBubbleColor;
  bubble.userBubbleColor = isValidHexColor(els.bubbleUserColorInput?.value) ? els.bubbleUserColorInput.value : defaults.customization.bubble.userBubbleColor;
  bubble.showAssistantName = Boolean(els.bubbleShowAssistantNameInput?.checked);
  bubble.showUserName = Boolean(els.bubbleShowUserNameInput?.checked);
  bubble.showAssistantAvatar = Boolean(els.bubbleShowAssistantAvatarInput?.checked);
  bubble.showUserAvatar = Boolean(els.bubbleShowUserAvatarInput?.checked);
  bubble.assistantAvatarSize = Math.min(160, Math.max(24, Number.parseInt(els.bubbleAssistantAvatarSizeInput?.value, 10) || state.customization.avatarSize));
  bubble.userAvatarSize = Math.min(160, Math.max(24, Number.parseInt(els.bubbleUserAvatarSizeInput?.value, 10) || state.customization.avatarSize));
  bubble.assistantAvatarShape = ["cutout", "circle", "square", "rectangle"].includes(els.bubbleAssistantAvatarShapeInput?.value)
    ? els.bubbleAssistantAvatarShapeInput.value
    : defaults.customization.bubble.assistantAvatarShape;
  bubble.userAvatarShape = ["cutout", "circle", "square", "rectangle"].includes(els.bubbleUserAvatarShapeInput?.value)
    ? els.bubbleUserAvatarShapeInput.value
    : defaults.customization.bubble.userAvatarShape;

  novel.assistantBackgroundColor = isValidHexColor(els.novelAssistantColorInput?.value) ? els.novelAssistantColorInput.value : defaults.customization.novel.assistantBackgroundColor;
  novel.userBackgroundColor = isValidHexColor(els.novelUserColorInput?.value) ? els.novelUserColorInput.value : defaults.customization.novel.userBackgroundColor;
  novel.chatBackgroundColor = isValidHexColor(els.novelChatBackgroundColorInput?.value) ? els.novelChatBackgroundColorInput.value : defaults.customization.novel.chatBackgroundColor;
  novel.assistantBackgroundBlur = Math.min(32, Math.max(0, Number.parseInt(els.novelAssistantBlurInput?.value, 10) || 0));
  novel.userBackgroundBlur = Math.min(32, Math.max(0, Number.parseInt(els.novelUserBlurInput?.value, 10) || 0));
  novel.chatBackgroundBlur = Math.min(32, Math.max(0, Number.parseInt(els.novelChatBackgroundBlurInput?.value, 10) || 0));
  novel.showAssistantName = Boolean(els.novelShowAssistantNameInput?.checked);
  novel.showUserName = Boolean(els.novelShowUserNameInput?.checked);
  novel.showAssistantAvatar = Boolean(els.novelShowAssistantAvatarInput?.checked);
  novel.showUserAvatar = Boolean(els.novelShowUserAvatarInput?.checked);
  novel.assistantAvatarSize = Math.min(200, Math.max(24, Number.parseInt(els.novelAssistantAvatarSizeInput?.value, 10) || defaults.customization.novel.assistantAvatarSize));
  novel.userAvatarSize = Math.min(200, Math.max(24, Number.parseInt(els.novelUserAvatarSizeInput?.value, 10) || defaults.customization.novel.userAvatarSize));
  novel.assistantAvatarShape = ["cutout", "circle", "square", "rectangle"].includes(els.novelAssistantAvatarShapeInput?.value)
    ? els.novelAssistantAvatarShapeInput.value
    : defaults.customization.novel.assistantAvatarShape;
  novel.userAvatarShape = ["cutout", "circle", "square", "rectangle"].includes(els.novelUserAvatarShapeInput?.value)
    ? els.novelUserAvatarShapeInput.value
    : defaults.customization.novel.userAvatarShape;
  novel.assistantAvatarAlign = ["left", "center", "right"].includes(els.novelAssistantAvatarAlignInput?.value)
    ? els.novelAssistantAvatarAlignInput.value
    : defaults.customization.novel.assistantAvatarAlign;
  novel.userAvatarAlign = ["left", "center", "right"].includes(els.novelUserAvatarAlignInput?.value)
    ? els.novelUserAvatarAlignInput.value
    : defaults.customization.novel.userAvatarAlign;
  novel.assistantTextFlow = ["wrap", "below"].includes(els.novelAssistantTextFlowInput?.value)
    ? els.novelAssistantTextFlowInput.value
    : defaults.customization.novel.assistantTextFlow;
  novel.userTextFlow = ["wrap", "below"].includes(els.novelUserTextFlowInput?.value)
    ? els.novelUserTextFlowInput.value
    : defaults.customization.novel.userTextFlow;

  state.customization.bubble = bubble;
  state.customization.novel = novel;
  return true;
}

function saveCustomization() {
  if (!collectCustomizationFromInputs()) {
    showToast("Customisation controls are unavailable", "error");
    return;
  }

  persistState();
  syncCustomizationInputs();
  applyCustomization();
  renderMessages();
  showToast("Customisation saved!", "success");
  setStatus("Customisation saved");
}

function resetCustomization() {
  state.customization = structuredClone(defaults.customization);
  persistState();
  syncCustomizationInputs();
  applyCustomization();
  renderMessages();
  showToast("Customisation reset", "success");
  setStatus("Customisation reset");
}

function onDefaultTextColorLiveChange() {
  if (!els.defaultTextColorInput) return;
  state.customization.defaultTextColor = els.defaultTextColorInput.value || defaults.customization.defaultTextColor;
  applyCustomization();
  renderMessages();
  persistState();
}

function onWrapRuleColorInputChange() {
  if (!els.wrapRuleColorInput || !els.wrapRuleHexInput) return;
  const color = els.wrapRuleColorInput.value;
  els.wrapRuleHexInput.value = color;
  const ruleId = els.wrapRuleEditor?.dataset?.ruleId;
  if (!ruleId) return;
  const rule = state.customization.wrapRules.find((item) => item.id === ruleId);
  if (!rule) return;
  rule.color = color;
  applyCustomization();
  renderMessages();
  persistState();
}

function onWrapRuleHexInputChange() {
  if (!els.wrapRuleHexInput || !els.wrapRuleColorInput) return;
  const value = els.wrapRuleHexInput.value.trim();
  if (!isValidHexColor(value)) return;
  els.wrapRuleColorInput.value = value;
  const ruleId = els.wrapRuleEditor?.dataset?.ruleId;
  if (!ruleId) return;
  const rule = state.customization.wrapRules.find((item) => item.id === ruleId);
  if (!rule) return;
  rule.color = value;
  applyCustomization();
  renderMessages();
  persistState();
}

function applyWrapRuleCssVariables() {
  state.customization.wrapRules.forEach((rule, index) => {
    document.documentElement.style.setProperty(`--chat-wrap-color-${index}`, rule.color || "#c084fc");
  });
}

function syncChatStyleSectionVisibility() {
  const style = state.customization.chatStyle === "novel" ? "novel" : "bubble";
  if (els.bubbleStyleSection) {
    els.bubbleStyleSection.classList.toggle("hidden", style !== "bubble");
  }
  if (els.novelStyleSection) {
    els.novelStyleSection.classList.toggle("hidden", style !== "novel");
  }
}

function onChatStyleSelectChange() {
  state.customization.chatStyle = els.chatStyleSelect?.value === "novel" ? "novel" : "bubble";
  syncChatStyleSectionVisibility();
  if (state.customization.chatStyle === "novel") {
    expandedMessageMenuIndex = -1;
    messageInlineEditIndex = -1;
  }
  applyCustomization();
  renderMessages();
}

async function onNovelImageSelected(event, target) {
  const [file] = event?.target?.files || [];
  if (!file) return;
  try {
    const dataUrl = await readFileAsDataUrl(file);
    if (target === "assistant") {
      state.customization.novel.assistantBackgroundImage = dataUrl;
    } else if (target === "user") {
      state.customization.novel.userBackgroundImage = dataUrl;
    } else {
      state.customization.novel.chatBackgroundImage = dataUrl;
    }
    persistState();
    syncCustomizationInputs();
    applyCustomization();
    renderMessages();
    showToast("Background image applied", "success");
  } catch {
    showToast("Failed to load image", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function openTextFontManager() {
  if (els.chatCustomizationSection?.classList.contains("hidden")) {
    els.chatCustomizationSection.classList.remove("hidden");
  }
  if (els.textFontSection?.classList.contains("hidden")) {
    els.textFontSection.classList.remove("hidden");
  }
  syncCustomizationInputs();
  toggleFontManager(true);
}

function toggleFontManager(show) {
  if (!els.fontManagerBackdrop) return;
  els.fontManagerBackdrop.classList.toggle("hidden", !show);
  if (show) {
    setInlineStatus(els.fontManagerStatus, "");
    syncCustomizationInputs();
  }
}

function isSupportedFontFile(fileName = "") {
  return /\.(ttf|otf|woff2?|TTF|OTF|WOFF2?)$/.test(String(fileName));
}

function getFontFormatFromFileName(fileName = "") {
  const lower = String(fileName).toLowerCase();
  if (lower.endsWith(".ttf")) return "truetype";
  if (lower.endsWith(".otf")) return "opentype";
  if (lower.endsWith(".woff")) return "woff";
  if (lower.endsWith(".woff2")) return "woff2";
  return "";
}

function extractFontFamilyFromFileName(fileName = "") {
  return String(fileName)
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[_-]+/g, " ")
    .trim();
}

function toSafeFontFamily(fontName = "") {
  const cleaned = String(fontName || "")
    .replace(/["'`]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || "Custom App Font";
}

function ensureAppFontStyleTag() {
  const existing = document.getElementById("appFontDynamicStyle");
  if (existing) {
    appFontStyleTag = existing;
    return existing;
  }
  const styleTag = document.createElement("style");
  styleTag.id = "appFontDynamicStyle";
  document.head.appendChild(styleTag);
  appFontStyleTag = styleTag;
  return styleTag;
}

function buildGoogleFontsCssUrl(fontName = "") {
  const token = String(fontName || "")
    .trim()
    .split(/\s+/)
    .map((part) => encodeURIComponent(part))
    .join("+");
  return `https://fonts.googleapis.com/css2?family=${token}&display=swap`;
}

function escapeCssValue(value) {
  return String(value || "")
    .replaceAll("\\", "\\\\")
    .replaceAll('"', '\\"')
    .replaceAll("\n", " ");
}

function applyAppFont() {
  const fallbackFontStack = "Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, \"Segoe UI\", sans-serif";
  const source = state.customization.appFontSource;
  let fontFamily = "";
  let fontCss = "";

  if (source === "file" && state.customization.appFontDataUrl) {
    fontFamily = toSafeFontFamily(state.customization.appFontFamily || extractFontFamilyFromFileName(state.customization.appFontFileName));
    const format = getFontFormatFromFileName(state.customization.appFontFileName);
    fontCss = `@font-face { font-family: "${escapeCssValue(fontFamily)}"; src: url("${escapeCssValue(state.customization.appFontDataUrl)}")${format ? ` format("${format}")` : ""}; font-display: swap; }`;
  } else if (source === "google" && (state.customization.appFontGoogleCss || state.customization.appFontGoogleName || state.customization.appFontFamily)) {
    fontFamily = toSafeFontFamily(state.customization.appFontFamily || state.customization.appFontGoogleName);
    if (state.customization.appFontGoogleCss) {
      fontCss = state.customization.appFontGoogleCss;
    } else if (state.customization.appFontGoogleName) {
      fontCss = `@import url("${escapeCssValue(buildGoogleFontsCssUrl(state.customization.appFontGoogleName))}");`;
    }
  }

  if (fontFamily && fontCss) {
    const styleTag = ensureAppFontStyleTag();
    styleTag.textContent = `${fontCss}\n:root { --app-font-family: "${escapeCssValue(fontFamily)}", ${fallbackFontStack}; }`;
    document.documentElement.style.setProperty("--app-font-family", `"${fontFamily}", ${fallbackFontStack}`);
  } else {
    document.documentElement.style.setProperty("--app-font-family", fallbackFontStack);
    if (appFontStyleTag) {
      appFontStyleTag.remove();
      appFontStyleTag = null;
    } else {
      document.getElementById("appFontDynamicStyle")?.remove();
    }
  }
}

async function onFontFileSelected(event) {
  const [file] = event?.target?.files || [];
  if (!file) return;
  if (!isSupportedFontFile(file.name)) {
    setInlineStatus(els.fontManagerStatus, "Unsupported file. Use .ttf, .otf, .woff, or .woff2.");
    showToast("Unsupported font file", "error");
    if (event?.target) event.target.value = "";
    return;
  }

  try {
    setInlineStatus(els.fontManagerStatus, "Loading font file...");
    const dataUrl = await readFileAsDataUrl(file);
    const family = toSafeFontFamily(extractFontFamilyFromFileName(file.name));

    state.customization.appFontSource = "file";
    state.customization.appFontFamily = family;
    state.customization.appFontFileName = file.name;
    state.customization.appFontDataUrl = dataUrl;
    state.customization.appFontGoogleName = "";
    state.customization.appFontGoogleCss = "";

    persistState();
    applyCustomization();
    syncCustomizationInputs();
    setInlineStatus(els.fontManagerStatus, `Applied \"${family}\" from file.`);
    showToast("Font applied", "success");
  } catch (error) {
    setInlineStatus(els.fontManagerStatus, error?.message || "Failed to read font file");
    showToast("Failed to load font file", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

async function addFontFromGoogleFonts() {
  const rawName = els.googleFontSearchInput?.value?.trim() || "";
  if (!rawName) {
    setInlineStatus(els.fontManagerStatus, "Enter a font name to search Google Fonts.");
    showToast("Enter a Google Font name", "error");
    return;
  }

  try {
    setInlineStatus(els.fontManagerStatus, `Checking Google Fonts for \"${rawName}\"...`);
    const url = buildGoogleFontsCssUrl(rawName);
    const response = await fetch(url, { method: "GET" });
    const css = await response.text();

    if (!response.ok || !css.includes("@font-face")) {
      throw new Error("No matching Google Font found.");
    }

    const familyMatch = css.match(/font-family:\s*['\"]([^'\"]+)['\"]/i);
    const resolvedFamily = toSafeFontFamily(familyMatch?.[1] || rawName);

    state.customization.appFontSource = "google";
    state.customization.appFontFamily = resolvedFamily;
    state.customization.appFontGoogleName = rawName;
    state.customization.appFontGoogleCss = css;
    state.customization.appFontFileName = "";
    state.customization.appFontDataUrl = "";

    persistState();
    applyCustomization();
    syncCustomizationInputs();
    setInlineStatus(els.fontManagerStatus, `Added Google Font: ${resolvedFamily}`);
    showToast("Google Font added", "success");
  } catch (error) {
    setInlineStatus(els.fontManagerStatus, error?.message || "Could not load that Google Font.");
    showToast("Google Font not found", "error");
  }
}

function resetAppFont() {
  state.customization.appFontSource = defaults.customization.appFontSource;
  state.customization.appFontFamily = defaults.customization.appFontFamily;
  state.customization.appFontFileName = "";
  state.customization.appFontDataUrl = "";
  state.customization.appFontGoogleName = "";
  state.customization.appFontGoogleCss = "";
  persistState();
  applyCustomization();
  syncCustomizationInputs();
  setInlineStatus(els.fontManagerStatus, "Font reset to system default.");
  showToast("Font reset", "success");
}

function renderImpersonationOptions() {
  const current = els.impersonateSelect.value;
  els.impersonateSelect.innerHTML = '<option value="">Speak as yourself</option>';
  state.characters.forEach((character) => {
    const option = document.createElement("option");
    option.value = character.id;
    option.textContent = character.name;
    option.selected = character.id === current;
    els.impersonateSelect.appendChild(option);
  });
  updateImpersonationButton();
}

function toggleImpersonationPicker() {
  const shouldShow = els.impersonateSelect.classList.contains("hidden");
  els.impersonateSelect.classList.toggle("hidden", !shouldShow);
  els.impersonateToggleBtn.classList.toggle("active", shouldShow);
  if (shouldShow) {
    els.impersonateSelect.focus();
  }
}

function onImpersonationChange() {
  updateImpersonationButton();
  els.impersonateSelect.classList.add("hidden");
  els.impersonateToggleBtn.classList.remove("active");
}

function updateImpersonationButton() {
  const selectedId = els.impersonateSelect.value;
  const selectedCharacter = state.characters.find((character) => character.id === selectedId);
  els.impersonateToggleBtn.title = selectedCharacter ? `Speaking as ${selectedCharacter.name}` : "Speak as yourself";
}

function setActiveDrawerPanel(panelId = "root") {
  if (!els.chatDrawer) return;
  els.chatDrawer.dataset.panel = panelId;
  els.drawerPanels?.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.panelId === panelId);
  });

  if (panelId === "provider") {
    showDrawerProviderEditor(state.activeProviderId);
  } else if (panelId === "preset-editor") {
    showDrawerPresetEditor(state.activePresetId);
  } else if (panelId === "lorebook") {
    renderDrawerLorebooks();
  } else if (panelId === "log") {
    renderDrawerLogs();
  }
}

function openDrawerPanel(panelId = "root") {
  if (!panelId) return;
  if (drawerPanelStack[drawerPanelStack.length - 1] !== panelId) {
    drawerPanelStack.push(panelId);
  }
  setActiveDrawerPanel(panelId);
}

function closeDrawerPanel() {
  if (drawerPanelStack.length <= 1) {
    toggleChatDrawer(false);
    return;
  }
  drawerPanelStack.pop();
  setActiveDrawerPanel(drawerPanelStack[drawerPanelStack.length - 1] || "root");
}

function toggleChatDrawer(show) {
  toggleSettings(false);
  if (show && state.activeView !== "chatsView") {
    return;
  }
  els.chatDrawer?.classList.toggle("open", Boolean(show));
  els.chatDrawerBackdrop?.classList.toggle("hidden", !show);
  if (show) {
    drawerPanelStack = ["root"];
    setActiveDrawerPanel("root");
    renderDrawerState();
  } else {
    hideDrawerPresetEditor();
    hideDrawerProviderEditor();
    if (els.messageActionPopup && !els.messageActionPopup.classList.contains("hidden")) {
      closeMessageActionPopup();
    }
  }
}

function onChatTitleInput() {
  state.currentChatTitle = els.chatTitleInput?.value?.trim() || "New chat";
  const chat = getCurrentChat();
  if (chat) {
    chat.title = state.currentChatTitle;
    syncCurrentChatFromState();
  }
  persistState();
  renderRecentChats();
}

function getActivePreset() {
  return state.presets.find((preset) => preset.id === state.activePresetId) || state.presets[0] || null;
}

function renderDrawerState() {
  if (els.chatTitleInput) {
    els.chatTitleInput.value = state.currentChatTitle || "New chat";
  }
  const preset = getActivePreset();
  const provider = state.providers.find((item) => item.id === state.activeProviderId);
  if (els.currentPresetLabel) {
    const topP = Number(preset?.topP ?? 1).toFixed(2);
    const temp = Number(preset?.temperature ?? 0.7).toFixed(2);
    const max = Number(preset?.maxTokens ?? 300);
    els.currentPresetLabel.textContent = `Current preset: ${preset?.name || "Default preset"} • temp ${temp} • top_p ${topP} • max ${max}`;
  }
  if (els.currentPresetSummaryBtn) {
    const summaryTitle = els.currentPresetSummaryBtn.querySelector("strong");
    if (summaryTitle) summaryTitle.textContent = preset?.name || "Default preset";
  }
  if (els.currentProviderLabel) {
    const thinkingEnabled = provider?.showThinkingSummaries !== false ? "on" : "off";
    els.currentProviderLabel.textContent = `Current provider: ${provider?.name || "None selected"} • model ${state.selectedModel || "-"} • thinking ${thinkingEnabled}`;
  }
  if (els.drawerPresetDetails) {
    const presetDetails = preset
      ? {
        temperature: preset.temperature,
        topP: preset.topP,
        maxTokens: preset.maxTokens,
        stream: Boolean(preset.stream),
        includeSystemPrompt: preset.useSystemPrompt !== false,
      }
      : {};
    els.drawerPresetDetails.textContent = Object.keys(presetDetails).length ? JSON.stringify(presetDetails, null, 2) : "";
  }
  if (els.drawerProviderDetails) {
    const providerDetails = provider
      ? {
        baseUrl: provider.baseUrl,
        selectedModel: state.selectedModel,
        showThinking: provider.showThinkingSummaries !== false,
        includeThinkingInPrompt: Boolean(provider.includeThinkingInPrompt),
        thinkingStartTag: provider.thinkingStartTag || "<thinking>",
        thinkingEndTag: provider.thinkingEndTag || "</thinking>",
      }
      : {};
    els.drawerProviderDetails.textContent = Object.keys(providerDetails).length ? JSON.stringify(providerDetails, null, 2) : "";
  }
  renderDrawerPresets();
  renderDrawerProviders();
  renderDrawerLorebooks();
  renderDrawerLogs();
  updateImpersonationButton();
}

function getCurrentChatCharacter() {
  const chat = getCurrentChat();
  if (!chat?.characterId) return null;
  return state.characters.find((character) => character.id === chat.characterId) || null;
}

function readAvatarCandidate(value) {
  if (!value) return "";
  if (typeof value === "string") {
    const normalized = value.trim();
    if (!normalized || ["none", "null", "undefined"].includes(normalized.toLowerCase())) return "";
    return normalized;
  }
  if (typeof value === "object") {
    const nested = value.dataUrl
      || value.dataURL
      || value.url
      || value.src
      || value.path
      || value.avatar;
    return readAvatarCandidate(nested);
  }
  return "";
}

function getCharacterAvatarUrl(character = null) {
  if (!character || typeof character !== "object") return "";
  return readAvatarCandidate(character.avatarDataUrl)
    || readAvatarCandidate(character.avatar)
    || readAvatarCandidate(character.avatarUrl)
    || readAvatarCandidate(character.avatar_url)
    || readAvatarCandidate(character.data?.avatarDataUrl)
    || readAvatarCandidate(character.data?.avatar)
    || readAvatarCandidate(character.data?.avatarUrl)
    || readAvatarCandidate(character.data?.avatar_url)
    || "";
}

function setElementBackgroundImage(element, rawUrl) {
  if (!element) return;
  const url = String(rawUrl || "").trim();
  if (!url) {
    element.style.backgroundImage = "";
    return;
  }
  element.style.backgroundImage = `url(${JSON.stringify(url)})`;
}

function getActiveChatStyle() {
  return state.customization?.chatStyle === "novel" ? "novel" : "bubble";
}

function shouldShowMessageAvatar(msg) {
  if (!msg || (msg.role !== "assistant" && msg.role !== "user")) return false;
  const style = getActiveChatStyle();
  const custom = style === "novel" ? state.customization.novel : state.customization.bubble;
  if (msg.role === "assistant") return custom?.showAssistantAvatar !== false;
  if (msg.role === "user") return custom?.showUserAvatar !== false;
  return false;
}

function getMessageAvatarSize(msg) {
  if (!msg || (msg.role !== "assistant" && msg.role !== "user")) return state.customization.avatarSize || 40;
  const style = getActiveChatStyle();
  const custom = style === "novel" ? state.customization.novel : state.customization.bubble;
  if (msg.role === "assistant") return Number(custom?.assistantAvatarSize || state.customization.avatarSize || 40);
  if (msg.role === "user") return Number(custom?.userAvatarSize || state.customization.avatarSize || 40);
  return state.customization.avatarSize || 40;
}

function getMessageAvatarShape(msg) {
  if (!msg || (msg.role !== "assistant" && msg.role !== "user")) return "cutout";
  const style = getActiveChatStyle();
  const custom = style === "novel" ? state.customization.novel : state.customization.bubble;
  if (msg.role === "assistant") return custom?.assistantAvatarShape || "cutout";
  if (msg.role === "user") return custom?.userAvatarShape || "cutout";
  return "cutout";
}

function getMessageAvatarAlign(msg) {
  if (!msg || (msg.role !== "assistant" && msg.role !== "user")) return "left";
  const style = getActiveChatStyle();
  if (style !== "novel") {
    return msg.role === "user" ? "right" : "left";
  }
  const custom = state.customization.novel || {};
  if (msg.role === "assistant") return custom.assistantAvatarAlign || "left";
  if (msg.role === "user") return custom.userAvatarAlign || "left";
  return "left";
}

function getMessageTextFlow(msg) {
  if (!msg || (msg.role !== "assistant" && msg.role !== "user")) return "wrap";
  const style = getActiveChatStyle();
  if (style !== "novel") return "wrap";
  const custom = state.customization.novel || {};
  if (msg.role === "assistant") return custom.assistantTextFlow || "wrap";
  if (msg.role === "user") return custom.userTextFlow || "wrap";
  return "wrap";
}

function getMessageAvatar(msg) {
  if (msg.role === "assistant") {
    return getCharacterAvatarUrl(getCurrentChatCharacter()) || readAvatarCandidate(msg?.avatarDataUrl);
  }
  if (msg.role === "user") {
    const selectedCharacter = state.characters.find((character) => character.id === els.impersonateSelect?.value);
    return getCharacterAvatarUrl(selectedCharacter);
  }
  return "";
}

function getMessageDisplayRole(msg) {
  if (msg.role === "assistant") {
    const style = getActiveChatStyle();
    const show = style === "novel"
      ? state.customization.novel?.showAssistantName !== false
      : state.customization.bubble?.showAssistantName !== false;
    if (!show) return "";
    return getCurrentChatCharacter()?.name || "assistant";
  }
  if (msg.role === "user") {
    const style = getActiveChatStyle();
    const show = style === "novel"
      ? state.customization.novel?.showUserName !== false
      : state.customization.bubble?.showUserName !== false;
    if (!show) return "";
    const selectedCharacter = state.characters.find((character) => character.id === els.impersonateSelect?.value);
    return selectedCharacter?.name || "user";
  }
  return msg.role;
}

function formatMessageContent(content) {
  let html = escapeHtml(content || "");
  state.customization.wrapRules.forEach((rule, index) => {
    if (rule?.enabled === false) return;

    const startRaw = String(rule.start || "");
    const endRaw = String(rule.end || "");
    if (!startRaw || !endRaw) return;

    const escapedStart = escapeHtml(startRaw);
    const escapedEnd = escapeHtml(endRaw);
    const start = escapeRegex(escapedStart);
    const end = escapeRegex(escapedEnd);
    if (!start || !end) return;

    const className = `format-wrap-${index}`;
    if (startRaw === "*" && endRaw === "*") {
      html = html.replace(new RegExp(`${start}([^\n]+?)${end}`, "g"), `<span class="${className}" style="color: var(--chat-wrap-color-${index})">$1</span>`);
      return;
    }

    html = html.replace(new RegExp(`${start}([^]+?)${end}`, "g"), `<span class="${className}" style="color: var(--chat-wrap-color-${index})">${escapedStart}$1${escapedEnd}</span>`);
  });
  return html;
}

function resolveCurrentUserName() {
  const selectedCharacter = state.characters.find((character) => character.id === els.impersonateSelect?.value);
  return selectedCharacter?.name?.trim() || "User";
}

function getMacroLocalVarsStore() {
  const chatId = state.currentChatId || "";
  if (!chatId) return {};
  if (!state.macroChatVars[chatId] || typeof state.macroChatVars[chatId] !== "object") {
    state.macroChatVars[chatId] = {};
  }
  return state.macroChatVars[chatId];
}

function normalizeMacroValue(value) {
  if (value === undefined || value === null) return "";
  return String(value);
}

function parseDiceFormula(formula) {
  const raw = String(formula || "").trim();
  if (!raw) return null;
  const normalized = /^\d+$/.test(raw) ? `1d${raw}` : raw;
  const match = normalized.match(/^(\d*)d(\d+)([+-]\d+)?$/i);
  if (!match) return null;
  const count = Number(match[1] || 1);
  const sides = Number(match[2]);
  const mod = Number(match[3] || 0);
  if (!Number.isFinite(count) || !Number.isFinite(sides) || !Number.isFinite(mod)) return null;
  if (count < 1 || sides < 1 || count > 1000 || sides > 100000) return null;
  return { count, sides, mod };
}

function rollDice(formula) {
  const parsed = parseDiceFormula(formula);
  if (!parsed) return "";
  let total = parsed.mod;
  for (let i = 0; i < parsed.count; i += 1) {
    total += 1 + Math.floor(Math.random() * parsed.sides);
  }
  return String(total);
}

function simpleHash(value) {
  let hash = 0;
  const str = String(value || "");
  for (let i = 0; i < str.length; i += 1) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return Math.abs(hash);
}

function splitMacroArgs(raw = "") {
  const input = String(raw || "");
  const parts = [];
  let current = "";
  for (let i = 0; i < input.length; i += 1) {
    const char = input[i];
    const next = input[i + 1];
    if (char === "\\" && next) {
      current += next;
      i += 1;
      continue;
    }
    if (char === ":" && next === ":") {
      parts.push(current.trim());
      current = "";
      i += 1;
      continue;
    }
    current += char;
  }
  parts.push(current.trim());
  return parts;
}

function splitMacroListArg(raw) {
  const value = String(raw || "");
  if (value.includes("::")) {
    return splitMacroArgs(value).map((item) => item.trim()).filter(Boolean);
  }
  const parts = [];
  let current = "";
  for (let i = 0; i < value.length; i += 1) {
    const char = value[i];
    const next = value[i + 1];
    if (char === "\\" && next) {
      current += next;
      i += 1;
      continue;
    }
    if (char === ",") {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  parts.push(current.trim());
  return parts.filter(Boolean);
}


function getLastMessageByRole(role = "") {
  const list = Array.isArray(state.messages) ? state.messages : [];
  for (let i = list.length - 1; i >= 0; i -= 1) {
    const msg = list[i];
    if (!msg || (msg.role !== "user" && msg.role !== "assistant")) continue;
    if (!role || msg.role === role) return msg;
  }
  return null;
}

function evaluateMacroCondition(rawCondition, context) {
  const resolved = applyMacrosToText(String(rawCondition || ""), context.character, context.userName, {
    vars: context.vars,
    skipScopedIf: true,
  }).trim();
  if (!resolved) return false;
  const lowered = resolved.toLowerCase();
  if (["0", "false", "off", "no", "null", "undefined"].includes(lowered)) return false;
  return true;
}

function applyMacrosToText(text, character = null, userNameOverride = "", options = {}) {
  const sourceText = String(text || "");
  const charRef = character || getCurrentChatCharacter();
  const charName = charRef?.name?.trim() || "Character";
  const userName = userNameOverride || resolveCurrentUserName();
  const lastMessage = getLastMessageByRole("");
  const lastUserMessage = getLastMessageByRole("user");
  const lastCharMessage = getLastMessageByRole("assistant");
  const preset = getActivePreset();
  const instruct = preset?.instruct || {};
  const systemPrompt = preset?.systemPrompt?.trim() || state.systemPrompt || "";
  const localVars = getMacroLocalVarsStore();
  const globalVars = state.macroGlobalVars || {};
  const creatorNotes = charRef?.data?.creator_notes || charRef?.data?.creatorNotes || "";
  const charVersion = charRef?.data?.character_version || charRef?.data?.char_version || charRef?.data?.version || "";
  const mesExamplesRaw = charRef?.data?.mes_example || charRef?.data?.mesExamples || "";
  const hasMessages = Array.isArray(state.messages) && state.messages.length > 0;
  const latestSwipeMessage = hasMessages
    ? [...state.messages].reverse().find((msg) => msg?.role === "assistant" && Array.isArray(msg.swipes) && msg.swipes.length)
    : null;
  const currentSwipeId = latestSwipeMessage
    ? clampSwipeIndex(Number(latestSwipeMessage.swipe_id ?? 0), latestSwipeMessage.swipes.length)
    : 0;
  const lastSwipeId = latestSwipeMessage ? Math.max(0, latestSwipeMessage.swipes.length - 1) : 0;
  const isMobile = typeof window !== "undefined" && typeof window.matchMedia === "function"
    ? window.matchMedia("(max-width: 768px)").matches
    : false;

  const baseVars = {
    char: charName,
    user: userName,
    charIfNotGroup: charName,
    group: charName,
    groupNotMuted: charName,
    notChar: userName,
    persona: "",
    charPrompt: "",
    charInstruction: "",
    charDescription: charRef?.data?.description || charRef?.personality || "",
    description: charRef?.data?.description || charRef?.personality || "",
    charPersonality: charRef?.personality || charRef?.data?.personality || "",
    personality: charRef?.personality || charRef?.data?.personality || "",
    charScenario: charRef?.scenario || charRef?.data?.scenario || "",
    scenario: charRef?.scenario || charRef?.data?.scenario || "",
    charFirstMessage: charRef?.greeting || charRef?.data?.first_mes || "",
    greeting: charRef?.greeting || charRef?.data?.first_mes || "",
    charCreatorNotes: creatorNotes,
    creatorNotes,
    charVersion,
    version: charVersion,
    char_version: charVersion,
    mesExamplesRaw,
    mesExamples: mesExamplesRaw,
    charDepthPrompt: "",
    original: charRef?.data?.original || "",
    model: state.selectedModel || "",
    systemPrompt,
    defaultSystemPrompt: systemPrompt,
    instructSystem: systemPrompt,
    instructSystemPrompt: systemPrompt,
    chatStart: preset?.chatStart || "",
    exampleSeparator: preset?.exampleSeparator || "",
    chatSeparator: preset?.exampleSeparator || "",
    maxPrompt: String(preset?.maxTokens ?? ""),
    maxPromptTokens: String(preset?.maxTokens ?? ""),
    maxContext: String(preset?.maxTokens ?? ""),
    maxContextTokens: String(preset?.maxTokens ?? ""),
    maxResponse: String(preset?.maxTokens ?? ""),
    maxResponseTokens: String(preset?.maxTokens ?? ""),
    isMobile: isMobile ? "true" : "false",
    lastGenerationType: String(state.lastGenerationType || ""),
    lastMessage: lastMessage?.content || "",
    lastUserMessage: lastUserMessage?.content || "",
    lastCharMessage: lastCharMessage?.content || "",
    lastMessageId: lastMessage ? String(Math.max(0, state.messages.length - 1)) : "",
    firstIncludedMessageId: hasMessages ? "0" : "",
    firstDisplayedMessageId: hasMessages ? "0" : "",
    currentSwipeId: String(currentSwipeId),
    lastSwipeId: String(lastSwipeId),
    allChatRange: state.messages?.length ? `0-${Math.max(0, state.messages.length - 1)}` : "",
    instructStoryStringPrefix: instruct.storyStringPrefix || "",
    instructStoryStringSuffix: instruct.storyStringSuffix || "",
    instructUserPrefix: instruct.inputSequence || "",
    instructInput: instruct.inputSequence || "",
    instructUserSuffix: instruct.inputSuffix || "",
    instructAssistantPrefix: instruct.outputSequence || "",
    instructOutput: instruct.outputSequence || "",
    instructAssistantSuffix: instruct.outputSuffix || instruct.stopSequence || "",
    instructSeparator: instruct.outputSuffix || instruct.stopSequence || "",
    instructSystemPrefix: instruct.systemSequence || "",
    instructSystemSuffix: instruct.systemSuffix || "",
    instructFirstAssistantPrefix: instruct.firstOutputSequence || instruct.outputSequence || "",
    instructFirstOutputPrefix: instruct.firstOutputSequence || instruct.outputSequence || "",
    instructLastAssistantPrefix: instruct.lastOutputSequence || instruct.outputSequence || "",
    instructLastOutputPrefix: instruct.lastOutputSequence || instruct.outputSequence || "",
    instructStop: instruct.stopSequence || "",
    instructUserFiller: "",
    instructSystemInstructionPrefix: instruct.systemSequence || "",
    instructFirstUserPrefix: instruct.firstInputSequence || instruct.inputSequence || "",
    instructFirstInput: instruct.firstInputSequence || instruct.inputSequence || "",
    instructLastUserPrefix: instruct.lastInputSequence || instruct.inputSequence || "",
    instructLastInput: instruct.lastInputSequence || instruct.inputSequence || "",
  };

  const vars = {
    ...baseVars,
    ...(options.vars && typeof options.vars === "object" ? options.vars : {}),
  };

  let output = sourceText;

  const resolveNamedMacro = (nameRaw, argRaw, fullToken = "") => {
    const macroName = String(nameRaw || "").trim();
    const macro = macroName.toLowerCase();
    const arg = String(argRaw || "");

    if (macro === "space") {
      const [count] = splitMacroArgs(arg);
      return " ".repeat(Math.max(1, Number(count || 1)));
    }
    if (macro === "newline") {
      const [count] = splitMacroArgs(arg);
      return "\n".repeat(Math.max(1, Number(count || 1)));
    }
    if (["noop", "trim", "banned", "outlet"].includes(macro)) return "";
    if (macro === "//" || macro === "comment") return "";
    if (macro === "datetimeformat") {
      const now = new Date();
      const format = arg.trim();
      if (!format) return now.toLocaleString();
      return format
        .replace(/YYYY/g, String(now.getFullYear()))
        .replace(/MM/g, String(now.getMonth() + 1).padStart(2, "0"))
        .replace(/DD/g, String(now.getDate()).padStart(2, "0"))
        .replace(/HH/g, String(now.getHours()).padStart(2, "0"))
        .replace(/mm/g, String(now.getMinutes()).padStart(2, "0"))
        .replace(/ss/g, String(now.getSeconds()).padStart(2, "0"));
    }
    if (macro === "time") return new Date().toLocaleTimeString();
    if (macro === "date") return new Date().toLocaleDateString();
    if (macro === "weekday") return new Date().toLocaleDateString(undefined, { weekday: "long" });
    if (macro === "isotime") return new Date().toISOString().slice(11, 16);
    if (macro === "isodate") return new Date().toISOString().slice(0, 10);
    if (macro === "idleduration" || macro === "idle_duration") return "just now";
    if (macro === "timediff") {
      const [left = "", right = ""] = splitMacroArgs(arg);
      const a = new Date(left.trim());
      const b = new Date(right.trim());
      if (Number.isNaN(a.getTime()) || Number.isNaN(b.getTime())) return "";
      const ms = Math.abs(a.getTime() - b.getTime());
      const mins = Math.floor(ms / 60000);
      if (mins < 60) return `${mins} minute${mins === 1 ? "" : "s"}`;
      const hours = Math.floor(mins / 60);
      if (hours < 24) return `${hours} hour${hours === 1 ? "" : "s"}`;
      const days = Math.floor(hours / 24);
      return `${days} day${days === 1 ? "" : "s"}`;
    }

    if ([
      "setvar", "addvar", "incvar", "decvar", "getvar", "hasvar", "varexists", "deletevar", "flushvar",
      "setglobalvar", "addglobalvar", "incglobalvar", "decglobalvar", "getglobalvar", "hasglobalvar", "globalvarexists", "deleteglobalvar", "flushglobalvar",
    ].includes(macro)) {
      const parts = [macroName, ...splitMacroArgs(arg)];
      const normalized = parts[0].toLowerCase();
      const name = String(parts[1] || "").trim();
      const rawValue = parts.slice(2).join("::").trim();
      const target = normalized.includes("global") ? globalVars : localVars;
      if (!name && !["incvar", "decvar", "incglobalvar", "decglobalvar"].includes(normalized)) return "";

      const current = target[name];
      const currentNum = Number(current);
      const valueNum = Number(rawValue);
      const canAddNum = Number.isFinite(currentNum) && Number.isFinite(valueNum);

      if (normalized === "setvar" || normalized === "setglobalvar") {
        target[name] = rawValue;
        return "";
      }
      if (normalized === "addvar" || normalized === "addglobalvar") {
        target[name] = canAddNum ? String(currentNum + valueNum) : `${normalizeMacroValue(current)}${rawValue}`;
        return "";
      }
      if (normalized === "incvar" || normalized === "incglobalvar") {
        target[name] = String((Number.isFinite(currentNum) ? currentNum : 0) + 1);
        return target[name];
      }
      if (normalized === "decvar" || normalized === "decglobalvar") {
        target[name] = String((Number.isFinite(currentNum) ? currentNum : 0) - 1);
        return target[name];
      }
      if (normalized === "getvar" || normalized === "getglobalvar") return normalizeMacroValue(target[name]);
      if (["hasvar", "varexists", "hasglobalvar", "globalvarexists"].includes(normalized)) {
        return Object.prototype.hasOwnProperty.call(target, name) ? "true" : "false";
      }
      if (["deletevar", "flushvar", "deleteglobalvar", "flushglobalvar"].includes(normalized)) {
        delete target[name];
        return "";
      }
      return "";
    }

    if (macro === "roll") return rollDice(arg);
    if (macro === "reverse") return Array.from(arg).reverse().join("");
    if (macro === "input") return String(els.messageInput?.value || "");
    if (macro === "hasextension") {
      const [extensionName = ""] = splitMacroArgs(arg);
      return extensionName ? "false" : "false";
    }
    if (macro === "random" || macro === "pick") {
      const items = splitMacroListArg(arg);
      if (!items.length) return "";
      if (macro === "pick") {
        const seed = simpleHash(`${state.currentChatId || ""}::${fullToken || `${macroName}::${arg}`}`);
        return items[seed % items.length] || "";
      }
      return items[Math.floor(Math.random() * items.length)] || "";
    }

    if (macro === "if") {
      const [cond = "", yesVal = "", noVal = ""] = splitMacroArgs(arg);
      return evaluateMacroCondition(cond, { character: charRef, userName, vars }) ? String(yesVal) : String(noVal);
    }

    if (Object.prototype.hasOwnProperty.call(vars, macroName)) {
      return String(vars[macroName] ?? "");
    }

    const key = Object.keys(vars).find((item) => item.toLowerCase() === macro);
    if (key) return String(vars[key] ?? "");

    return "";
  };

  if (!options.skipScopedIf) {
    output = output.replace(/{{\s*if\s+([^}]+)}}([\s\S]*?){{\s*\/if\s*}}/gi, (_, conditionRaw, body) => {
      const [thenBranch, elseBranch = ""] = String(body || "").split(/{{\s*else\s*}}/i);
      return evaluateMacroCondition(conditionRaw, { character: charRef, userName, vars }) ? thenBranch : elseBranch;
    });
  }

  let changedInLastPass = false;
  for (let pass = 0; pass < 16; pass += 1) {
    const before = output;
    output = output.replace(/{{\s*([^{}]+?)\s*}}/g, (full, innerRaw) => {
      const inner = String(innerRaw || "").trim();
      if (!inner) return "";
      if (/^#if\s+/i.test(inner) || /^\/if$/i.test(inner) || /^else$/i.test(inner)) return "";
      const [name, ...rest] = splitMacroArgs(inner);
      const arg = rest.join("::");
      return resolveNamedMacro(name, arg, full);
    });

    output = output
      .replace(/<USER>/gi, userName)
      .replace(/<BOT>/gi, charName)
      .replace(/<CHAR>/gi, charName);

    changedInLastPass = before !== output;
    if (!changedInLastPass) break;
  }

  if (changedInLastPass && /{{\s*[^{}]+\s*}}/.test(output)) {
    console.warn("MaiTavern macro engine: reached max-pass limit while resolving macros", {
      sample: output.slice(0, 240),
    });
  }

  persistState();
  return output;
}

function parseAlternateGreetings(value, options = {}) {
  const fromEditor = Boolean(options?.fromEditor);

  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }

  const text = String(value || "").replace(/\r\n/g, "\n").trim();
  if (!text) return [];

  if (/\n\s*---\s*\n/.test(text)) {
    return text
      .split(/\n\s*---\s*\n/g)
      .map((item) => item.trim())
      .filter(Boolean);
  }

  if (!fromEditor && /^\s*\[[\s\S]*\]\s*$/.test(text)) {
    try {
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        return parsed.map((item) => String(item || "").trim()).filter(Boolean);
      }
    } catch {
      // ignore and continue with legacy parsing
    }
  }

  if (fromEditor) {
    return [text];
  }

  const lines = text.split("\n");
  const hasBlankLines = lines.some((line) => !line.trim());
  if (hasBlankLines) {
    return [text];
  }

  return lines.map((item) => item.trim()).filter(Boolean);
}

function formatAlternateGreetingsForEditor(value) {
  const greetings = parseAlternateGreetings(value);
  if (!greetings.length) return "";
  return greetings.join("\n\n---\n\n");
}

function getCharacterGreetingVariants(character) {
  const primary = String(character?.greeting || character?.data?.first_mes || "").trim();
  const alternates = parseAlternateGreetings(character?.alternateGreetings || character?.data?.alternate_greetings || []);
  return [primary, ...alternates].filter(Boolean);
}

function normalizeCharacterRecord(item = {}) {
  const data = item?.data && typeof item.data === "object" ? item.data : item;
  const name = data?.name || item?.name || "";
  const scenario = data?.scenario || item?.scenario || "";
  const personality = data?.personality || item?.personality || data?.description || item?.description || "";
  const greeting = data?.first_mes || item?.first_mes || data?.greeting || item?.greeting || "";
  const alternateGreetings = parseAlternateGreetings(
    item?.alternateGreetings
    || data?.alternate_greetings
    || data?.alternateGreetings
    || item?.alternate_greetings
    || ""
  );
  const avatarDataUrl = readAvatarCandidate(
    item?.avatarDataUrl
    || item?.avatar
    || item?.avatarUrl
    || item?.avatar_url
    || data?.avatarDataUrl
    || data?.avatar
    || data?.avatarUrl
    || data?.avatar_url
    || ""
  );
  const fallbackDescription = name || "Imported character";
  return {
    id: item?.id || crypto.randomUUID(),
    name,
    scenario,
    personality,
    greeting,
    alternateGreetings,
    description: personality || scenario || item?.description || fallbackDescription,
    avatarDataUrl,
    data: {
      ...(data && typeof data === "object" ? data : {}),
      name,
      scenario,
      personality,
      first_mes: greeting,
      alternate_greetings: alternateGreetings,
      description: personality || scenario || item?.description || "",
      avatarDataUrl,
      avatar: avatarDataUrl,
    },
  };
}

function normalizeChatRecord(chat = {}) {
  const lorebookSettingsSource = chat?.lorebookSettings && typeof chat.lorebookSettings === "object"
    ? chat.lorebookSettings
    : defaults.lorebookSettings;

  return {
    ...chat,
    lorebookIds: Array.isArray(chat.lorebookIds)
      ? chat.lorebookIds.map((item) => String(item || "").trim()).filter(Boolean)
      : [],
    lorebookSettings: {
      scanDepth: Math.min(64, Math.max(1, Number.parseInt(lorebookSettingsSource?.scanDepth, 10) || defaults.lorebookSettings.scanDepth)),
      tokenBudget: Math.min(8000, Math.max(64, Number.parseInt(lorebookSettingsSource?.tokenBudget, 10) || defaults.lorebookSettings.tokenBudget)),
      recursiveScanning: lorebookSettingsSource?.recursiveScanning !== false,
    },
    messages: Array.isArray(chat.messages) ? chat.messages.map((message) => normalizeMessageRecord(message)) : [],
  };
}

function normalizeMessageRecord(message = {}) {
  const normalized = { ...message };
  if (Array.isArray(normalized.swipes)) {
    normalized.swipes = normalized.swipes.map((item) => String(item ?? ""));
    normalized.swipe_id = clampSwipeIndex(Number(normalized.swipe_id ?? 0), normalized.swipes.length);
    normalized.content = normalized.swipes[normalized.swipe_id] || normalized.content || "";
  }
  return normalized;
}

function clampSwipeIndex(index, length) {
  if (!Number.isFinite(index)) return 0;
  if (!length || length < 1) return 0;
  return Math.min(Math.max(0, index), length - 1);
}

function getMessageSwipeContent(message) {
  if (!Array.isArray(message?.swipes) || !message.swipes.length) {
    return message?.content || "";
  }
  const index = clampSwipeIndex(Number(message.swipe_id ?? 0), message.swipes.length);
  return message.swipes[index] || message.content || "";
}

function setMessageSwipeIndex(messageIndex, nextIndex) {
  const message = state.messages[messageIndex];
  if (!message || !Array.isArray(message.swipes) || message.swipes.length < 2) return;
  const clamped = clampSwipeIndex(nextIndex, message.swipes.length);
  message.swipe_id = clamped;
  message.content = message.swipes[clamped] || message.content || "";
  syncCurrentChatFromState();
  persistState();
  renderMessages();
  renderRecentChats();
}

function createUserFilePayload() {
  return {
    app: "maitavern",
    schemaVersion: USER_FILE_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    state: JSON.parse(JSON.stringify(state)),
  };
}

function hasFileSystemAccessApi() {
  return typeof fetch === "function";
}

function canUseBrowserStorage() {
  try {
    if (typeof localStorage === "undefined") return false;
    const probeKey = "__maitavern_probe__";
    localStorage.setItem(probeKey, "1");
    localStorage.removeItem(probeKey);
    return true;
  } catch {
    return false;
  }
}

function canUseIndexedDb() {
  return typeof indexedDB !== "undefined";
}

function getUserFileIndexedDb() {
  if (!canUseIndexedDb()) {
    return Promise.reject(new Error("IndexedDB unavailable"));
  }

  if (!userFileIndexedDbPromise) {
    userFileIndexedDbPromise = new Promise((resolve, reject) => {
      let request;
      try {
        request = indexedDB.open(USER_FILE_INDEXED_DB_NAME, 1);
      } catch (error) {
        reject(error);
        return;
      }

      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(USER_FILE_INDEXED_DB_STORE)) {
          db.createObjectStore(USER_FILE_INDEXED_DB_STORE);
        }
      };

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onerror = () => {
        reject(request.error || new Error("IndexedDB open failed"));
      };
    });
  }

  return userFileIndexedDbPromise;
}

async function writeUserFileToIndexedDb(payload) {
  if (!canUseIndexedDb()) return false;

  try {
    const db = await getUserFileIndexedDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(USER_FILE_INDEXED_DB_STORE, "readwrite");
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("IndexedDB write transaction failed"));
      tx.onabort = () => reject(tx.error || new Error("IndexedDB write transaction aborted"));
      tx.objectStore(USER_FILE_INDEXED_DB_STORE).put(payload, USER_FILE_INDEXED_DB_KEY);
    });
    return true;
  } catch {
    return false;
  }
}

async function readUserFileFromIndexedDb() {
  if (!canUseIndexedDb()) return null;

  try {
    const db = await getUserFileIndexedDb();
    const value = await new Promise((resolve, reject) => {
      const tx = db.transaction(USER_FILE_INDEXED_DB_STORE, "readonly");
      tx.onerror = () => reject(tx.error || new Error("IndexedDB read transaction failed"));
      tx.onabort = () => reject(tx.error || new Error("IndexedDB read transaction aborted"));
      const request = tx.objectStore(USER_FILE_INDEXED_DB_STORE).get(USER_FILE_INDEXED_DB_KEY);
      request.onsuccess = () => resolve(request.result ?? null);
      request.onerror = () => reject(request.error || new Error("IndexedDB read failed"));
    });

    return value && typeof value === "object" ? value : null;
  } catch {
    return null;
  }
}

function writeUserFileToBrowserStorage(payload) {
  if (!canUseBrowserStorage()) return false;

  try {
    const text = JSON.stringify(payload);
    localStorage.setItem(USER_FILE_STORAGE_KEY, text);
  } catch {
    return false;
  }

  // Legacy key is best-effort only. Do not fail the save if this overflows quota.
  try {
    if (payload?.state && typeof payload.state === "object") {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload.state));
    }
  } catch {
    // no-op
  }

  return true;
}

function readUserFileFromBrowserStorage() {
  if (!canUseBrowserStorage()) return null;

  try {
    const userFileRaw = localStorage.getItem(USER_FILE_STORAGE_KEY);
    if (userFileRaw) {
      const parsed = JSON.parse(userFileRaw);
      if (parsed && typeof parsed === "object") {
        return parsed;
      }
    }
  } catch {
    // keep trying legacy key
  }

  try {
    const legacyRaw = localStorage.getItem(STORAGE_KEY);
    if (!legacyRaw) return null;
    const parsedState = JSON.parse(legacyRaw);
    if (!parsedState || typeof parsedState !== "object") return null;
    return {
      app: "maitavern",
      schemaVersion: USER_FILE_SCHEMA_VERSION,
      exportedAt: new Date().toISOString(),
      state: parsedState,
    };
  } catch {
    return null;
  }
}

async function writeUserFileToDisk(payload) {
  try {
    const response = await fetch(USER_FILE_ENDPOINT, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(text || `HTTP ${response.status}`);
    }

    deviceStorageUnavailable = false;
  } catch (error) {
    deviceStorageUnavailable = true;
    const indexedDbSaved = await writeUserFileToIndexedDb(payload);
    if (indexedDbSaved) {
      writeUserFileToBrowserStorage(payload);
      return;
    }
    if (writeUserFileToBrowserStorage(payload)) {
      return;
    }
    throw new Error(`Local device storage unavailable. Start MaiTavern server (${error?.message || "request failed"})`);
  }
}

async function readUserFileFromDisk() {
  try {
    const response = await fetch(USER_FILE_ENDPOINT, {
      method: "GET",
      headers: { "Accept": "application/json" },
    });

    if (response.status === 404) {
      throw new Error("User file not found");
    }

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(text || `HTTP ${response.status}`);
    }

    const parsed = await response.json();
    if (!parsed || typeof parsed !== "object") {
      throw new Error("User file is invalid");
    }

    deviceStorageUnavailable = false;
    return { parsed };
  } catch (error) {
    const indexedDbFallback = await readUserFileFromIndexedDb();
    if (indexedDbFallback && typeof indexedDbFallback === "object") {
      deviceStorageUnavailable = true;
      return { parsed: indexedDbFallback };
    }

    const fallback = readUserFileFromBrowserStorage();
    if (fallback && typeof fallback === "object") {
      deviceStorageUnavailable = true;
      return { parsed: fallback };
    }

    deviceStorageUnavailable = true;
    throw new Error(`Local device storage unavailable. Start MaiTavern server (${error?.message || "request failed"})`);
  }
}

async function flushPendingUserFileWrite() {
  if (!pendingUserFilePayload || userFileWriteInFlight) return;

  const payload = pendingUserFilePayload;
  pendingUserFilePayload = null;
  userFileWriteInFlight = true;

  try {
    await writeUserFileToDisk(payload);
    setStatus(deviceStorageUnavailable ? "Saved locally on device" : `Saved to local file: ${USER_FILE_NAME}`);
  } catch {
    // Preserve payload for retry unless a newer one is already queued.
    if (!pendingUserFilePayload) {
      pendingUserFilePayload = payload;
    }
  } finally {
    userFileWriteInFlight = false;
    if (pendingUserFilePayload) {
      if (userFileWriteTimer) clearTimeout(userFileWriteTimer);
      userFileWriteTimer = setTimeout(() => {
        void flushPendingUserFileWrite();
      }, 300);
    }
  }
}

function persistUserFile() {
  const payload = createUserFilePayload();

  if (!hasFileSystemAccessApi()) {
    writeUserFileToBrowserStorage(payload);
    return;
  }

  pendingUserFilePayload = payload;
  if (userFileWriteTimer) {
    clearTimeout(userFileWriteTimer);
  }
  userFileWriteTimer = setTimeout(() => {
    void flushPendingUserFileWrite();
  }, 220);
}

function loadUserFile() {
  return null;
}

function loadLegacyStateFromStorage() {
  const storedPayload = readUserFileFromBrowserStorage();
  const importedState = storedPayload?.state && typeof storedPayload.state === "object"
    ? storedPayload.state
    : storedPayload && typeof storedPayload === "object"
      ? storedPayload
      : null;

  if (importedState) {
    return {
      ...structuredClone(defaults),
      ...importedState,
    };
  }

  return structuredClone(defaults);
}

async function exportUserDataZip() {
  try {
    const payloadText = JSON.stringify(createUserFilePayload(), null, 2);
    const zipBytes = createSingleFileZip("maitavern-user.json", payloadText);
    downloadBlob("maitavern-data.zip", new Blob([zipBytes], { type: "application/zip" }));
    setDataStatus("Export completed");
    setStatus("Data exported");
    showToast("Data exported", "success");
  } catch (error) {
    setDataStatus(error?.message || "Export failed");
    setStatus("Data export failed");
    showToast("Data export failed", "error");
  }
}

async function importUserDataZip(event) {
  const [file] = event?.target?.files || [];
  if (!file) return;

  try {
    setDataStatus("Importing...");
    const buffer = await file.arrayBuffer();
    const zipEntry = extractSingleFileFromZip(buffer);
    const parsed = JSON.parse(zipEntry.text);
    const importedState = parsed?.state && typeof parsed.state === "object" ? parsed.state : parsed;

    state = {
      ...structuredClone(defaults),
      ...importedState,
    };

    ensureCollections();
    syncInputsFromState();
    renderModelOptions();
    applyCustomization();
    renderMessages();
    renderCharacters();
    renderProviders();
    renderPresets();
    renderLorebooks();
    renderRecentChats();
    renderImpersonationOptions();
    renderDrawerState();
    syncCustomizationInputs();
    syncLorebookSettingsInputs();
    renderWrapRuleList();
    applyWelcomeState();
    switchView(state.activeView || "homeView");
    updateChatUiVisibility();
    persistState();

    setDataStatus(`Imported ${zipEntry.name}`);
    setStatus("Data imported");
    showToast("Data imported", "success");
  } catch (error) {
    setDataStatus(error?.message || "Import failed");
    setStatus("Data import failed");
    showToast("Data import failed", "error");
  } finally {
    if (event?.target) event.target.value = "";
  }
}

function createSingleFileZip(filename, textContent) {
  const encoder = new TextEncoder();
  const nameBytes = encoder.encode(filename);
  const dataBytes = encoder.encode(textContent);
  const crc = crc32(dataBytes);

  const localHeader = new Uint8Array(30 + nameBytes.length);
  const localView = new DataView(localHeader.buffer);
  localView.setUint32(0, 0x04034b50, true);
  localView.setUint16(4, 20, true);
  localView.setUint16(6, 0, true);
  localView.setUint16(8, 0, true);
  localView.setUint16(10, 0, true);
  localView.setUint16(12, 0, true);
  localView.setUint32(14, crc, true);
  localView.setUint32(18, dataBytes.length, true);
  localView.setUint32(22, dataBytes.length, true);
  localView.setUint16(26, nameBytes.length, true);
  localView.setUint16(28, 0, true);
  localHeader.set(nameBytes, 30);

  const centralHeader = new Uint8Array(46 + nameBytes.length);
  const centralView = new DataView(centralHeader.buffer);
  centralView.setUint32(0, 0x02014b50, true);
  centralView.setUint16(4, 20, true);
  centralView.setUint16(6, 20, true);
  centralView.setUint16(8, 0, true);
  centralView.setUint16(10, 0, true);
  centralView.setUint16(12, 0, true);
  centralView.setUint16(14, 0, true);
  centralView.setUint32(16, crc, true);
  centralView.setUint32(20, dataBytes.length, true);
  centralView.setUint32(24, dataBytes.length, true);
  centralView.setUint16(28, nameBytes.length, true);
  centralView.setUint16(30, 0, true);
  centralView.setUint16(32, 0, true);
  centralView.setUint16(34, 0, true);
  centralView.setUint16(36, 0, true);
  centralView.setUint32(38, 0, true);
  centralView.setUint32(42, 0, true);
  centralHeader.set(nameBytes, 46);

  const endHeader = new Uint8Array(22);
  const endView = new DataView(endHeader.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, 1, true);
  endView.setUint16(10, 1, true);
  endView.setUint32(12, centralHeader.length, true);
  endView.setUint32(16, localHeader.length + dataBytes.length, true);
  endView.setUint16(20, 0, true);

  return concatUint8Arrays([localHeader, dataBytes, centralHeader, endHeader]);
}

function extractSingleFileFromZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);

  const signatureOffset = findZipLocalHeaderOffset(bytes);
  if (signatureOffset < 0) {
    throw new Error("Invalid ZIP file");
  }

  const signature = view.getUint32(signatureOffset, true);
  if (signature !== 0x04034b50) {
    throw new Error("ZIP local file header not found");
  }

  const compressionMethod = view.getUint16(signatureOffset + 8, true);
  if (compressionMethod !== 0) {
    throw new Error("Unsupported ZIP compression method");
  }

  const compressedSize = view.getUint32(signatureOffset + 18, true);
  const fileNameLength = view.getUint16(signatureOffset + 26, true);
  const extraLength = view.getUint16(signatureOffset + 28, true);

  const nameStart = signatureOffset + 30;
  const dataStart = nameStart + fileNameLength + extraLength;
  const dataEnd = dataStart + compressedSize;

  if (dataEnd > bytes.length) {
    throw new Error("Corrupted ZIP file data");
  }

  const decoder = new TextDecoder();
  const name = decoder.decode(bytes.slice(nameStart, nameStart + fileNameLength));
  const text = decoder.decode(bytes.slice(dataStart, dataEnd));
  return { name, text };
}

function findZipLocalHeaderOffset(bytes) {
  for (let i = 0; i < bytes.length - 3; i += 1) {
    if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) {
      return i;
    }
  }
  return -1;
}

function concatUint8Arrays(arrays) {
  const total = arrays.reduce((sum, arr) => sum + arr.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  arrays.forEach((arr) => {
    out.set(arr, offset);
    offset += arr.length;
  });
  return out;
}

let crcTable = null;

function getCrcTable() {
  if (crcTable) return crcTable;
  crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    crcTable[i] = c >>> 0;
  }
  return crcTable;
}

function crc32(bytes) {
  const table = getCrcTable();
  let crc = 0 ^ (-1);
  for (let i = 0; i < bytes.length; i += 1) {
    crc = (crc >>> 8) ^ table[(crc ^ bytes[i]) & 0xff];
  }
  return (crc ^ (-1)) >>> 0;
}

async function loadState() {
  // Prefer local device file first.
  if (hasFileSystemAccessApi()) {
    try {
      await pickAndLoadUserFile();
      return state;
    } catch {
      // User may cancel picker; fall back safely.
    }
  }

  return loadLegacyStateFromStorage();
}

function migrateLegacyChats() {
  if (state.chats.length) {
    if (state.currentChatId && !state.chats.some((chat) => chat.id === state.currentChatId)) {
      state.currentChatId = state.chats[0]?.id || "";
    }
    if (!state.currentChatId && state.chats[0]) {
      state.currentChatId = state.chats[0].id;
      state.currentChatTitle = state.chats[0].title || "New chat";
      state.messages = Array.isArray(state.chats[0].messages) ? state.chats[0].messages.map((message) => ({ ...message })) : [];
    }
    return;
  }

  const usefulMessages = Array.isArray(state.messages)
    ? state.messages.filter((message) => message.role === "user" || message.role === "assistant")
    : [];

  if (!usefulMessages.length && !state.recentChats.length) {
    state.messages = [];
    return;
  }

  const migratedChat = {
    id: crypto.randomUUID(),
    title: state.currentChatTitle || state.recentChats[0]?.title || "New chat",
    messages: usefulMessages.map((message) => ({ ...message })),
    preview: (state.recentChats[0]?.preview || usefulMessages[usefulMessages.length - 1]?.content || "").slice(0, 80),
    updatedAt: state.recentChats[0]?.updatedAt || new Date().toISOString(),
  };

  state.chats = [migratedChat];
  state.recentChats = [{
    id: migratedChat.id,
    title: migratedChat.title,
    preview: migratedChat.preview,
    updatedAt: migratedChat.updatedAt,
  }];
  state.currentChatId = migratedChat.id;
  state.currentChatTitle = migratedChat.title;
  state.messages = migratedChat.messages.map((message) => ({ ...message }));
}

function autoResizeTextarea(textarea) {
  textarea.style.height = "auto";
  textarea.style.height = Math.min(textarea.scrollHeight, 180) + "px";
}

function downloadJson(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  downloadBlob(filename, blob);
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function slugify(value) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "item";
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
    reader.readAsArrayBuffer(file);
  });
}

function decodeBase64Unicode(base64Value) {
  const normalized = String(base64Value || "").trim().replace(/-/g, "+").replace(/_/g, "/");
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new TextDecoder().decode(bytes);
}

function normalizePngChunkKeyword(rawKeyword = "") {
  return String(rawKeyword || "").replace(/\0+$/g, "").trim().toLowerCase();
}

function parseEmbeddedCharacterJson(rawText = "") {
  const text = String(rawText || "").trim();
  if (!text) return null;

  const directCandidates = [text];
  if (text.startsWith("\ufeff")) {
    directCandidates.push(text.replace(/^\ufeff/, ""));
  }

  for (const candidate of directCandidates) {
    try {
      const parsed = JSON.parse(candidate);
      return parsed?.data || parsed;
    } catch {
      // continue
    }
  }

  try {
    const decoded = decodeBase64Unicode(text);
    const parsed = JSON.parse(decoded);
    return parsed?.data || parsed;
  } catch {
    // continue
  }

  return null;
}

function normalizeImportedCharacter(item, fileName = "") {
  const normalized = normalizeCharacterRecord(item || {});
  if (!normalized.name) {
    normalized.name = fileName.replace(/\.json$/i, "") || "Imported character";
    normalized.data.name = normalized.name;
  }
  if (!normalized.description) {
    normalized.description = normalized.name || "Imported character";
  }
  return {
    ...normalized,
    id: crypto.randomUUID(),
  };
}

async function importCharacterCardFromImage(file) {
  const avatarDataUrl = await readFileAsDataUrl(file);
  const embedded = await extractCharacterCardFromPng(file).catch(() => null);
  if (embedded) {
    const normalized = normalizeImportedCharacter(embedded, file.name);
    normalized.avatarDataUrl = avatarDataUrl;
    if (normalized.data && typeof normalized.data === "object") {
      normalized.data.avatarDataUrl = avatarDataUrl;
    }
    return normalized;
  }
  const name = file.name.replace(/\.[^.]+$/, "") || "Imported character";
  return {
    id: crypto.randomUUID(),
    name,
    description: "Imported from avatar image",
    avatarDataUrl,
    data: {
      name,
      description: "Imported from avatar image",
      avatarDataUrl,
      spec: "chara_card_v2",
    },
  };
}

async function extractCharacterCardFromPng(file) {
  const buffer = await readFileAsArrayBuffer(file);
  const bytes = new Uint8Array(buffer);

  if (bytes.length < 8) {
    throw new Error("PNG file is too small");
  }

  const pngSignature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  const isPng = pngSignature.every((value, index) => bytes[index] === value);
  if (!isPng) {
    throw new Error("File is not a valid PNG");
  }

  const decoder = new TextDecoder();
  const view = new DataView(buffer);
  let offset = 8;
  const textCandidates = [];

  while (offset + 8 <= bytes.length) {
    const length = view.getUint32(offset, false);
    const typeStart = offset + 4;
    const typeEnd = typeStart + 4;
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    const crcEnd = dataEnd + 4;

    if (crcEnd > bytes.length) break;

    const chunkType = decoder.decode(bytes.slice(typeStart, typeEnd));
    const chunkData = bytes.slice(dataStart, dataEnd);

    if (chunkType === "tEXt") {
      const separatorIndex = chunkData.indexOf(0);
      if (separatorIndex > 0) {
        const keyword = normalizePngChunkKeyword(decoder.decode(chunkData.slice(0, separatorIndex)));
        const textValue = decoder.decode(chunkData.slice(separatorIndex + 1));
        textCandidates.push({ keyword, textValue });
      }
    } else if (chunkType === "iTXt") {
      const null1 = chunkData.indexOf(0);
      if (null1 > 0) {
        const keyword = normalizePngChunkKeyword(decoder.decode(chunkData.slice(0, null1)));
        const compressionFlag = chunkData[null1 + 1] ?? 0;
        const compressionMethod = chunkData[null1 + 2] ?? 0;
        let cursor = null1 + 3;
        const null2 = chunkData.indexOf(0, cursor);
        if (null2 >= 0) {
          cursor = null2 + 1;
          const null3 = chunkData.indexOf(0, cursor);
          if (null3 >= 0) {
            cursor = null3 + 1;
            if (compressionFlag === 0 && compressionMethod === 0) {
              const textValue = decoder.decode(chunkData.slice(cursor));
              textCandidates.push({ keyword, textValue });
            }
          }
        }
      }
    }

    if (chunkType === "IEND") break;
    offset = crcEnd;
  }

  const prioritized = [
    ...textCandidates.filter((item) => item.keyword === "chara"),
    ...textCandidates.filter((item) => item.keyword !== "chara"),
  ];

  for (const candidate of prioritized) {
    const parsed = parseEmbeddedCharacterJson(candidate.textValue);
    if (parsed && typeof parsed === "object") {
      return parsed;
    }
  }

  throw new Error("No embedded character data found in image");
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;
  return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
}

function isValidHexColor(value) {
  return /^#(?:[0-9a-fA-F]{3}){1,2}$/.test(String(value || "").trim());
}

function escapeAttribute(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("'", "&#39;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
