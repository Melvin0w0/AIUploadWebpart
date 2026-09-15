import * as React from 'react';
import {
  DatePicker,
  DayOfWeek,
  DefaultButton,
  defaultDatePickerStrings,
  Dropdown,
  IconButton,
  IDropdownOption,
  Link,
  PrimaryButton,
  ProgressIndicator,
  TextField
} from '@fluentui/react';
import styles from './AiUpload.module.scss';
import type { IAiUploadProps } from './IAiUploadProps';
import * as strings from 'AiUploadWebPartStrings';
import { locFormat } from '../loc/locFormat';
import { PdfOcrService } from '../services/PdfOcrService';
import { IOcrPageResult, IOcrProgress } from '../services/IPdfOcr';
import { buildOcrFieldMarks, formatOcrTextWithDebugMarks, IOcrFieldMark } from '../services/ocrFieldMarks';
import { formatOcrTextWithStyles, joinOcrWords, stripOcrStyleTags } from '../services/ocrSelection';
import PdfHighlightViewer from './PdfHighlightViewer';
import { DEFAULT_FORM_FIELDS, isNameField, isOrganizationField, isReceiverField, isRefNoField, isReadOnlyFormField, isRegistrationNumberField, isRequiredField, isSenderField, isSubjectField, missingRequiredFields } from '../constants/defaultFormFields';
import { CorrespondenceKind, correspondenceKindFromFileName, generateIncomingName, isIncomingName, nameFromPdfFile } from '../constants/incomingName';
import {
  canonicalLeadingBl,
  isLeadingBlField,
  LEADING_BL_OPTIONS,
  resolveLeadingBlSite
} from '../constants/blSiteMap';
import {
  eoiProjectNumberFromCode,
  isEoiProjectNumber,
  isProjectNumberField,
  isValidProjectNumber,
  sanitizeProjectNumber
} from '../constants/projectNumber';
import {
  canonicalSubProjectNumber,
  isSubProjectNumberField,
  SUB_PROJECT_NONE,
  SUB_PROJECT_NUMBER_OPTIONS
} from '../constants/subProjectNumber';
import {
  formatIssueDate,
  isIssueDateField,
  ISSUE_DATE_DISPLAY_LABEL,
  parseIssueDate,
  sanitizeIssueDate
} from '../constants/issueDate';
import {
  canonicalYesNo,
  isAttachmentField,
  isCcToAecomField,
  isScanField,
  isYesNoChoiceField,
  NO_VALUE,
  YES_NO_OPTIONS,
  YES_VALUE
} from '../constants/yesNo';
import {
  canonicalUploadTypeForProjectNumber,
  constrainUploadType,
  UploadType,
  UPLOAD_TYPE_NORMAL,
  uploadTypeOptionsForProjectNumber
} from '../constants/uploadType';
import {
  canonicalLabelType,
  LabelType,
  LABEL_TYPE_NORMAL,
  LABEL_TYPE_OPTIONS
} from '../constants/labelType';
import { extractFieldValues } from '../services/fieldExtractor';
import { extractFieldsWithAi, isAiExtractionConfigured } from '../services/AiFieldExtractor';
import { detectIncomingFields, incomingAiHints, pickIncomingFieldValue } from '../services/incoming/fields';
import { createIncomingOcrDecider, keepIncomingEmailPreviewPages } from '../services/incoming/email';
import { detectOutgoingFields, outgoingAiHints, pickOutgoingFieldValue } from '../services/outgoing/fields';
import { SharePointUploadService } from '../services/SharePointUploadService';
import {
  buildUploadFolderUrl,
  fileNameFromFields,
  resolveUploadDestination
} from '../services/uploadDestination';
import { correspondenceFolderName, lookupProjectUploadTypes } from '../services/projectUploadFolders';
import {
  rememberFieldValue,
  rememberFieldValues,
  isHistoryTextField,
  loadFieldHistory,
  saveFieldHistory,
  suggestionsFor,
  IFieldHistory
} from '../services/fieldHistory';
import { lookupLabelStaffFromNotificationSetup, lookupLeadingBlFromNotificationSetup, lookupNotificationSetupByProjectName, INotificationSetupLookup, emptyLabelStaff } from '../services/notificationSetup';
import { lookupRootUrlMappingCode, ROOT_URL_MAPPING_LIST_TITLE } from '../services/rootUrlMapping';
import { generateLabelPagePng } from '../services/labelPage';
import { appendLabelPageToPdf } from '../services/pdfLabelAppend';
import { isDevToolsOpen, isSpfxServeDebug, subscribeDevToolsOpen } from '../services/spfxLocalDebug';

interface IFormField {
  id: string;
  label: string;
  value: string;
  debugSource?: string;
}

interface IAiUploadState {
  file: File | undefined;
  pages: IOcrPageResult[];
  currentPage: number;
  selectedWordIndexes: number[];
  fields: IFormField[];
  activeFieldId: string | undefined;
  isProcessing: boolean;
  progress: IOcrProgress | undefined;
  error: string | undefined;
  info: string | undefined;
  success: string | undefined;
  successUrl: string | undefined;
  successFolderUrl: string | undefined;
  warning: string | undefined;
  isUploading: boolean;
  uploadStatus: string | undefined;
  showRequiredErrors: boolean;
  showOcrStyles: boolean;
  isRestyling: boolean;
  devToolsOpen: boolean;
  fieldDebugMarks: IOcrFieldMark[];
  history: IFieldHistory;
  historyFieldId: string | undefined;
  uploadType: UploadType;
  labelType: LabelType;
  incomingLetterType: string;
  dearToSubjectText: string;
  projectUploadTypes: UploadType[];
}

interface IChoiceOption {
  key: string;
  text: string;
}

interface IChoiceGroupProps {
  label: string;
  selected: string;
  options: IChoiceOption[];
  onChange: (key: string) => void;
  disabled: boolean;
  className?: string;
}

interface IChoiceThumb {
  x: number;
  y: number;
  width: number;
  height: number;
}

function AppleChoiceGroup(props: IChoiceGroupProps): React.ReactElement {
  const trackRef = React.useRef<HTMLDivElement>(null);
  const selectedRef = React.useRef<HTMLButtonElement>(null);
  const [thumb, setThumb] = React.useState<IChoiceThumb>({ x: 0, y: 0, width: 0, height: 0 });
  const [thumbReady, setThumbReady] = React.useState<boolean>(false);
  const optionSignature = props.options.map((option) => option.key + ':' + option.text).join('|');

  const syncThumb = React.useCallback((): void => {
    const track = trackRef.current;
    const button = selectedRef.current;
    if (!track || !button) {
      return;
    }
    const trackBox = track.getBoundingClientRect();
    const buttonBox = button.getBoundingClientRect();
    setThumb({
      x: buttonBox.left - trackBox.left,
      y: buttonBox.top - trackBox.top,
      width: buttonBox.width,
      height: buttonBox.height
    });
  }, []);

  React.useLayoutEffect(() => {
    syncThumb();
    const frame = window.requestAnimationFrame(() => {
      setThumbReady(true);
    });
    return () => {
      window.cancelAnimationFrame(frame);
    };
  }, [props.selected, optionSignature, syncThumb]);

  React.useEffect(() => {
    const track = trackRef.current;
    if (!track) {
      return;
    }
    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(() => {
        syncThumb();
      });
      observer.observe(track);
      return () => {
        observer.disconnect();
      };
    }
    const onResize = (): void => {
      syncThumb();
    };
    window.addEventListener('resize', onResize);
    return () => {
      window.removeEventListener('resize', onResize);
    };
  }, [syncThumb]);

  const pulseTrack = (target: HTMLElement): void => {
    const group = target.closest(`.${styles.choiceTrack}`) as HTMLElement | null;
    if (!group || typeof group.animate !== 'function') {
      return;
    }
    group.animate(
      [
        { transform: 'scale(1)' },
        { transform: 'scale(0.98)', offset: 0.32 },
        { transform: 'scale(1)' }
      ],
      { duration: 320, easing: 'cubic-bezier(0.32, 0.72, 0, 1)' }
    );
  };

  return (
    <div className={`${styles.choiceGroup} ${props.className || ''}`.trim()}>
      <div className={styles.choiceLabel}>{props.label}</div>
      <div ref={trackRef} className={styles.choiceTrack} role="radiogroup" aria-label={props.label}>
        <span
          className={`${styles.choiceThumb} ${thumbReady ? styles.choiceThumbReady : ''}`}
          style={{
            width: thumb.width,
            height: thumb.height,
            transform: `translate(${thumb.x}px, ${thumb.y}px)`
          }}
          aria-hidden={true}
        />
        {props.options.map((option) => {
          const active = option.key === props.selected;
          return (
            <button
              key={option.key}
              type="button"
              ref={active ? selectedRef : undefined}
              className={`${styles.choicePill} ${active ? styles.choicePillActive : ''}`}
              role="radio"
              aria-checked={active}
              disabled={props.disabled}
              onClick={(event) => {
                event.preventDefault();
                pulseTrack(event.currentTarget);
                props.onChange(option.key);
              }}
            >
              {option.text}
            </button>
          );
        })}
      </div>
    </div>
  );
}

export default class AiUpload extends React.Component<IAiUploadProps, IAiUploadState> {
  private _fileInput: React.RefObject<HTMLInputElement>;
  private _nextFieldId: number;
  private _originalPdfBytes: Uint8Array | undefined;
  private _calendarOpen: boolean;
  private _calendarObserver: MutationObserver | undefined;
  private _originalFocus: ((this: HTMLElement, options?: FocusOptions) => void) | undefined;
  private _focusPatchTimer: number | undefined;
  private _historyCloseTimer: number | undefined;
  private _leadingBlLookupSeq: number = 0;
  private _leadingBlLookupTimer: number | undefined;
  private _eoiLookupSeq: number = 0;
  private _eoiLookupTimer: number | undefined;
  private _uploadFolderLookupSeq: number = 0;
  private _uploadFolderLookupTimer: number | undefined;
  private _stopDevToolsWatch: (() => void) | undefined;
  private _restyleSeq: number;

  public constructor(props: IAiUploadProps) {
    super(props);
    this._fileInput = React.createRef<HTMLInputElement>();
    this._nextFieldId = 1;
    this._originalPdfBytes = undefined;
    this._calendarOpen = false;
    this._leadingBlLookupSeq = 0;
    this._leadingBlLookupTimer = undefined;
    this._eoiLookupSeq = 0;
    this._eoiLookupTimer = undefined;
    this._uploadFolderLookupSeq = 0;
    this._uploadFolderLookupTimer = undefined;
    this._stopDevToolsWatch = undefined;
    this._restyleSeq = 0;
    const fields = this._fieldsFromConfig(props.formFields);
    const devToolsOpen = isDevToolsOpen();
    this.state = {
      file: undefined,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      fields,
      activeFieldId: this._defaultActiveFieldId(fields),
      isProcessing: false,
      progress: undefined,
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      successFolderUrl: undefined,
      warning: undefined,
      isUploading: false,
      uploadStatus: undefined,
      showRequiredErrors: false,
      showOcrStyles: false,
      isRestyling: false,
      devToolsOpen,
      fieldDebugMarks: [],
      history: loadFieldHistory(),
      historyFieldId: undefined,
      uploadType: UPLOAD_TYPE_NORMAL,
      labelType: LABEL_TYPE_NORMAL,
      incomingLetterType: '',
      dearToSubjectText: '',
      projectUploadTypes: []
    };
  }

  public componentDidMount(): void {
    document.addEventListener('mousedown', this._onDocumentMouseDownCapture, true);
    document.addEventListener('click', this._onDocumentClickCapture, true);
    document.addEventListener('submit', this._onDocumentSubmitCapture, true);
    this._stopDevToolsWatch = subscribeDevToolsOpen(this._onDevToolsOpenChange);
  }

  public componentDidUpdate(prevProps: IAiUploadProps): void {
    if (prevProps.formFields !== this.props.formFields) {
      const fields = this._fieldsFromConfig(this.props.formFields, this.state.fields);
      this.setState({
        fields,
        activeFieldId: this._defaultActiveFieldId(fields)
      });
    }
  }

  public componentWillUnmount(): void {
    document.removeEventListener('mousedown', this._onDocumentMouseDownCapture, true);
    document.removeEventListener('click', this._onDocumentClickCapture, true);
    document.removeEventListener('submit', this._onDocumentSubmitCapture, true);
    this._stopCalendarObserver();
    this._disablePreventScrollFocus();
    if (this._historyCloseTimer !== undefined) {
      window.clearTimeout(this._historyCloseTimer);
    }
    if (this._leadingBlLookupTimer !== undefined) {
      window.clearTimeout(this._leadingBlLookupTimer);
    }
    if (this._eoiLookupTimer !== undefined) {
      window.clearTimeout(this._eoiLookupTimer);
    }
    if (this._uploadFolderLookupTimer !== undefined) {
      window.clearTimeout(this._uploadFolderLookupTimer);
    }
    this._leadingBlLookupSeq = this._leadingBlLookupSeq + 1;
    this._eoiLookupSeq = this._eoiLookupSeq + 1;
    this._uploadFolderLookupSeq = this._uploadFolderLookupSeq + 1;
    this._restyleSeq = this._restyleSeq + 1;
    this._revokePageUrls(this.state.pages);
    if (this._stopDevToolsWatch) {
      this._stopDevToolsWatch();
      this._stopDevToolsWatch = undefined;
    }
  }

  public render(): React.ReactElement<IAiUploadProps> {
    const { hasTeamsContext } = this.props;
    const {
      file,
      pages,
      currentPage,
      selectedWordIndexes,
      fields,
      activeFieldId,
      isProcessing,
      progress,
      error,
      info,
      success,
      successUrl,
      successFolderUrl,
      warning,
      isUploading,
      uploadStatus,
      showRequiredErrors,
      showOcrStyles,
      isRestyling,
      devToolsOpen,
      fieldDebugMarks,
      historyFieldId,
      uploadType,
      labelType,
      incomingLetterType,
      dearToSubjectText,
      projectUploadTypes
    } = this.state;
    const busy = isProcessing || isUploading || isRestyling;
    const converted = pages.length > 0 && !isProcessing;
    const requiredMissing = missingRequiredFields(fields);
    const markRequired = converted || showRequiredErrors;
    const percent = progress ? Math.max(0, Math.min(100, progress.percent)) / 100 : 0;
    const currentPreview = pages.length > 0 ? pages[Math.max(0, currentPage - 1)] : undefined;
    const pageFieldMarks = currentPreview
      ? fieldDebugMarks.filter((mark) => mark.pageNumber === currentPreview.pageNumber)
      : [];
    const ocrInspectText = currentPreview
      ? (
        showOcrStyles
          ? (formatOcrTextWithDebugMarks(currentPreview.words || [], pageFieldMarks, true) || currentPreview.text || '')
          : (joinOcrWords(currentPreview.words || []) || currentPreview.text || '')
      )
      : '';
    const hasFieldValues = fields.some((field) => field.value.length > 0);
    const documentKind = file ? correspondenceKindFromFileName(file.name) : 'unknown';
    const projectNumber = this._namedValue(fields, isProjectNumberField);
    const availableUploadTypes = projectUploadTypes.length > 0
      ? projectUploadTypes
      : uploadTypeOptionsForProjectNumber(projectNumber).map((option) => option.key);
    const resolvedUploadType = constrainUploadType(uploadType, availableUploadTypes);
    const destination = resolveUploadDestination(fields, {
      tenantUrl: this.props.tenantUrl,
      libraryName: this.props.libraryName,
      folderPathTemplate: this.props.folderPathTemplate,
      uploadType: resolvedUploadType,
      correspondenceKind: documentKind
    });
    const destinationUrl = buildUploadFolderUrl(destination);
    const destinationLabel = !destination.siteUrl
      ? strings.UploadDestinationPending
      : destinationUrl;
    const isLocalDebug = isSpfxServeDebug();
    const showDebugUi = isLocalDebug || devToolsOpen;

    return (
      <section className={`${styles.aiUpload} ${hasTeamsContext ? styles.teams : ''} ${showDebugUi ? styles.debugMode : ''}`}>
        {showDebugUi && (
          <div className={styles.debugBanner} role="status">
            {strings.ServeDebugBanner || 'DEBUG'}
          </div>
        )}
        <div className={styles.header}>
          <h1 className={styles.title}>{strings.WebPartTitle}</h1>
          <p className={styles.subtitle}>{strings.WebPartSubtitle}</p>
        </div>

        {error && this._renderBanner('error', error, this._clearError)}
        {!error && markRequired && requiredMissing.length > 0 && this._renderBanner(
          'error',
          locFormat(strings.RequiredFieldsPrompt, 'Please fill the required fields: {0}', requiredMissing.join(', '))
        )}
        {info && this._renderBanner('info', info, this._clearInfo)}
        {success && (
          <div className={`${styles.banner} ${styles.bannerSuccess}`} role="status">
            <div className={styles.bannerBody}>
              <div className={styles.bannerText}>{success}</div>
              {(successUrl || successFolderUrl) && (
                <div className={styles.bannerActions}>
                  {successUrl && (
                    <a
                      className={styles.appleBtn}
                      href={successUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      data-interception="off"
                      onClick={(event) => this._openInNewTab(event, successUrl)}
                    >
                      {strings.OpenUploadedFile || 'Open file'}
                    </a>
                  )}
                  {successFolderUrl && (
                    <a
                      className={`${styles.appleBtn} ${styles.appleBtnSecondary}`}
                      href={successFolderUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      data-interception="off"
                      onClick={(event) => this._openInNewTab(event, successFolderUrl)}
                    >
                      {strings.OpenUploadedFolder || 'Open folder'}
                    </a>
                  )}
                </div>
              )}
            </div>
            <button
              type="button"
              className={styles.bannerDismiss}
              onClick={this._clearSuccess}
              aria-label={strings.Dismiss}
            >
              ×
            </button>
          </div>
        )}
        {warning && this._renderBanner('warning', warning, this._clearWarning)}

        <div className={styles.toolbar}>
          <input
            ref={this._fileInput}
            type="file"
            accept="application/pdf,.pdf"
            className={styles.hiddenFileInput}
            onChange={this._onFileChange}
            disabled={busy}
          />
          <div className={styles.fileMeta}>
            <span className={styles.fileLabel}>{strings.SelectPdfLabel}</span>
            <span className={styles.fileNameRow}>
              <span className={styles.fileName}>{file ? file.name : strings.ChooseFile}</span>
              {documentKind !== 'unknown' && (
                <span className={`${styles.kindBadge} ${documentKind === 'incoming' ? styles.kindIncoming : styles.kindOutgoing}`}>
                  {documentKind === 'incoming'
                    ? (strings.IncomingLabel || 'Incoming')
                    : (strings.OutgoingLabel || 'Outgoing')}
                </span>
              )}
            </span>
          </div>
          <div className={styles.toolbarActions}>
            <DefaultButton
              className={styles.secondaryBtn}
              text={strings.ChooseFile}
              onClick={this._openFilePicker}
              disabled={busy}
            />
            <PrimaryButton
              className={styles.primaryBtn}
              text={strings.ConvertButton}
              onClick={this._onConvert}
              disabled={!file || busy}
            />
          </div>
        </div>

        {isProcessing && progress && (
          <div className={styles.progressWrap}>
            <ProgressIndicator
              label={progress.status}
              description={
                progress.totalPages > 0
                  ? locFormat(
                    strings.PageProgress,
                    'Page {0} of {1}',
                    String(progress.page),
                    String(progress.totalPages)
                  )
                  : undefined
              }
              percentComplete={percent}
            />
          </div>
        )}

          <div className={styles.results}>
          <div className={styles.pane}>
            <div className={styles.paneHeader}>
              <span className={styles.paneTitle}>{strings.FormFieldsLabel}</span>
            </div>
            <div className={styles.fieldsBody}>
              <p className={styles.hint}>{strings.HighlightHint}</p>
              {documentKind === 'incoming' && converted && incomingLetterType !== 'email' && (
                <div className={styles.dearSubjectBox}>
                  <div className={styles.dearSubjectLabel}>
                    {strings.IncomingDearToSubjectLabel || 'Below Dear, above Subject'}
                  </div>
                  <pre className={`${styles.dearSubjectText} ${dearToSubjectText ? '' : styles.dearSubjectEmpty}`.trim()}>
                    {dearToSubjectText || strings.IncomingDearToSubjectEmpty || 'No text found between Dear and Subject.'}
                  </pre>
                </div>
              )}
              <div className={styles.fieldGroup}>
                {fields.map((field) => (
                  <div
                    key={field.id}
                    className={`${styles.fieldCard} ${field.id === activeFieldId ? styles.fieldCardActive : ''} ${this._isMissingRequired(field, markRequired) ? styles.fieldCardMissing : ''} ${field.id === historyFieldId ? styles.fieldCardHistoryOpen : ''}`}
                    onClick={() => this._setActiveField(field.id)}
                  >
                    {isYesNoChoiceField(field.label) ? (
                      this._renderYesNo(field)
                    ) : (
                    <div className={styles.fieldRow}>
                      {this._renderFieldName(field, documentKind)}
                      {isLeadingBlField(field.label) ? (
                        <Dropdown
                          selectedKey={field.value || undefined}
                          options={this._leadingBlOptions(field.value)}
                          onChange={(_event, option) => this._onFieldValueChange(field.id, option ? String(option.key) : '')}
                          onFocus={() => this._setActiveField(field.id)}
                          placeholder={strings.LeadingBlPlaceholder}
                          ariaLabel={field.label}
                          errorMessage={this._requiredError(field, markRequired)}
                          className={styles.fieldInput}
                        />
                      ) : isSubProjectNumberField(field.label) ? (
                        <Dropdown
                          selectedKey={canonicalSubProjectNumber(field.value)}
                          options={this._subProjectNumberOptions()}
                          onChange={(_event, option) => this._onFieldValueChange(field.id, option ? String(option.key) : SUB_PROJECT_NONE)}
                          onFocus={() => this._setActiveField(field.id)}
                          ariaLabel={field.label}
                          errorMessage={this._requiredError(field, markRequired)}
                          className={styles.fieldInput}
                        />
                      ) : isIssueDateField(field.label) ? (
                        <div className={styles.fieldInput}>
                          <DatePicker
                            className={styles.datePicker}
                            value={parseIssueDate(field.value)}
                            onSelectDate={(date) => this._onIssueDateSelect(field.id, date)}
                            formatDate={(date) => date ? formatIssueDate(date) : ''}
                            parseDateFromString={(text) => parseIssueDate(text) || null}
                            placeholder={strings.IssueDatePlaceholder || 'dd/MM/yyyy'}
                            allowTextInput={true}
                            disableAutoFocus={true}
                            firstDayOfWeek={DayOfWeek.Monday}
                            strings={defaultDatePickerStrings}
                            ariaLabel={ISSUE_DATE_DISPLAY_LABEL}
                            isMonthPickerVisible={false}
                            calendarProps={{
                              showGoToToday: false,
                              styles: {
                                root: {
                                  width: 280
                                }
                              }
                            }}
                            calloutProps={{
                              className: styles.datePickerCallout,
                              gapSpace: 8,
                              isBeakVisible: false,
                              setInitialFocus: false,
                              preventDismissOnResize: true,
                              onMouseDown: (event) => this._patchCalendarButtons(event.currentTarget),
                              onClick: (event) => event.preventDefault(),
                              styles: {
                                root: {
                                  borderRadius: 16,
                                  overflow: 'hidden'
                                },
                                calloutMain: {
                                  borderRadius: 16,
                                  overflow: 'hidden'
                                }
                              },
                              layerProps: {
                                onLayerDidMount: this._onCalendarLayerMount,
                                onLayerWillUnmount: this._onCalendarLayerUnmount
                              }
                            }}
                            textField={{
                              borderless: true,
                              errorMessage: this._requiredError(field, markRequired),
                              onFocus: () => this._setActiveField(field.id),
                              onKeyDown: (event) => {
                                if (event.key === 'Enter') {
                                  event.preventDefault();
                                }
                              }
                            }}
                          />
                        </div>
                      ) : (
                        <div className={styles.historyField}>
                        <TextField
                          value={field.value}
                          onChange={(_event, newValue) => this._onFieldValueChange(field.id, newValue || '')}
                          onFocus={() => {
                            this._setActiveField(field.id);
                            if (!isReadOnlyFormField(field.label)) {
                              this._openHistory(field.id);
                            }
                          }}
                          onBlur={() => this._onHistoryFieldBlur(field)}
                          placeholder={
                            isNameField(field.label)
                              ? (documentKind === 'incoming'
                                ? (strings.IncomingNamePlaceholder || 'Generated Incoming number')
                                : strings.NamePlaceholder)
                              : isRegistrationNumberField(field.label)
                                ? strings.RegistrationNumberPlaceholder
                                : isProjectNumberField(field.label)
                                  ? strings.ProjectNumberPlaceholder
                                  : strings.FieldPlaceholder
                          }
                          ariaLabel={field.label}
                          maxLength={isProjectNumberField(field.label) ? 8 : undefined}
                          readOnly={isReadOnlyFormField(field.label)}
                          errorMessage={this._requiredError(field, markRequired)}
                          className={styles.fieldInput}
                          borderless={true}
                        />
                        {this._renderFieldHistory(field)}
                        </div>
                      )}
                      {!isReadOnlyFormField(field.label) && (
                      <IconButton
                        className={styles.clearFieldBtn}
                        iconProps={{ iconName: 'Cancel' }}
                        title={strings.ClearField}
                        ariaLabel={strings.ClearField}
                        onClick={(event) => {
                          event.stopPropagation();
                          this._onFieldValueChange(field.id, '');
                        }}
                      />
                      )}
                    </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          </div>

            <div className={styles.previewColumn}>
            <div className={styles.pane}>
              <div className={styles.paneHeader}>
                <span className={styles.paneTitle}>{strings.PdfPreviewLabel}</span>
                <div className={styles.previewHeaderActions}>
                  {showDebugUi && currentPreview && (
                    <button
                      type="button"
                      className={styles.debugBtn}
                      onClick={this._onRestylePage}
                      disabled={busy}
                    >
                      {isRestyling
                        ? (strings.RestylePageBusy || 'Restyling page…')
                        : (strings.RestylePageButton || 'Restyle page')}
                    </button>
                  )}
                  <div className={styles.pageNav}>
                  <IconButton
                    className={styles.pageBtn}
                    iconProps={{ iconName: 'ChevronLeft' }}
                    title={strings.PreviousPage}
                    ariaLabel={strings.PreviousPage}
                    disabled={pages.length === 0 || currentPage <= 1}
                    onClick={this._onPreviousPage}
                  />
                  <span className={styles.pageLabel}>
                    {locFormat(
                      strings.PageProgress,
                      'Page {0} of {1}',
                      String(pages.length > 0 ? currentPage : 0),
                      String(pages.length)
                    )}
                  </span>
                  <IconButton
                    className={styles.pageBtn}
                    iconProps={{ iconName: 'ChevronRight' }}
                    title={strings.NextPage}
                    ariaLabel={strings.NextPage}
                    disabled={pages.length === 0 || currentPage >= pages.length}
                    onClick={this._onNextPage}
                  />
                </div>
                </div>
              </div>
              {showOcrStyles && currentPreview && (
                <div className={styles.styleSpanLegend}>
                  <span className={styles.styleSpanLegendUnderline}>
                    {strings.StyleLegendUnderline || 'Underline'}
                  </span>
                  <span className={styles.styleSpanLegendSeparator}>
                    {strings.StyleLegendSeparator || 'Separator'}
                  </span>
                </div>
              )}
              {currentPreview ? (
                <PdfHighlightViewer
                  page={currentPreview}
                  selectedIndexes={selectedWordIndexes}
                  showStyles={showOcrStyles}
                  fieldMarks={pageFieldMarks}
                  onSelectText={this._onPdfSelectText}
                />
              ) : (
                <div className={styles.placeholder}>
                  {strings.PdfPreviewPlaceholder}
                </div>
              )}
            </div>

            <div className={styles.ocrTextPane}>
              <div className={styles.paneHeader}>
                <span className={styles.paneTitle}>{strings.ExtractedTextLabel}</span>
                <button
                  type="button"
                  className={`${styles.debugBtn} ${showOcrStyles ? styles.debugBtnOn : ''}`}
                  onClick={this._onToggleOcrStyles}
                >
                  Debug
                </button>
              </div>
              <div className={styles.ocrTextBody}>
                <p className={styles.hint}>
                  {showOcrStyles
                    ? (strings.DebugExtractedTextDescription || 'Debug: <b>/<u> are bold/underline. Field tags such as <Sender:署名括號> mark the OCR text used to fill each field.')
                    : strings.ExtractedTextDescription}
                </p>
                {showOcrStyles && pageFieldMarks.length > 0 && (
                  <div className={styles.fieldMarkLegend}>
                    {pageFieldMarks.map((mark) => (
                      <span key={mark.label + mark.pageNumber} className={styles.fieldMarkLegendItem} style={{ borderColor: mark.color, color: mark.color }}>
                        {mark.source ? (mark.label + ' · ' + mark.source) : mark.label}
                      </span>
                    ))}
                  </div>
                )}
                <TextField
                  multiline={true}
                  readOnly={true}
                  resizable={true}
                  rows={8}
                  value={ocrInspectText}
                  placeholder={strings.ExtractedTextPlaceholder}
                  className={styles.ocrTextField}
                  borderless={true}
                />
              </div>
            </div>
            </div>
          </div>

          <div className={styles.footer}>
            <div className={styles.uploadBar}>
              <p className={styles.hint}>{strings.UploadHint}</p>
              <div className={styles.uploadSelectors}>
                {this._renderChoiceGroup(
                  strings.UploadTypeLabel || 'Upload Type',
                  resolvedUploadType,
                  availableUploadTypes.map((key) => ({
                    key,
                    text: this._uploadTypeLabel(key)
                  })),
                  (key) => this._onUploadTypeChange(key),
                  busy,
                  styles.uploadType
                )}
                {documentKind === 'incoming' && this._renderChoiceGroup(
                  strings.LabelTypeLabel || 'Label Type',
                  labelType,
                  LABEL_TYPE_OPTIONS.map((option) => ({
                    key: option.key,
                    text: this._labelTypeLabel(option.key)
                  })),
                  (key) => this._onLabelTypeChange(key),
                  busy,
                  styles.labelType
                )}
              </div>
              <div className={styles.destination}>
                {strings.UploadDestinationLabel}:{' '}
                {destinationUrl && destination.siteUrl ? (
                  <Link
                    href={destinationUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    data-interception="off"
                    onClick={(event) => this._openInNewTab(event, destinationUrl)}
                  >
                    {destinationUrl}
                  </Link>
                ) : (
                  destinationLabel
                )}
              </div>
              {isUploading && uploadStatus && (
                <ProgressIndicator label={uploadStatus} />
              )}
            </div>
            <div className={styles.footerActions}>
              <DefaultButton
                className={styles.ghostBtn}
                text={strings.ClearButton}
                onClick={this._onClear}
                disabled={busy || (!file && pages.length === 0 && !hasFieldValues)}
              />
              {file && (
                <PrimaryButton
                  className={styles.primaryBtn}
                  text={strings.UploadButton}
                  onClick={this._onUpload}
                  disabled={busy}
                />
              )}
            </div>
          </div>
      </section>
    );
  }

  private _renderChoiceGroup = (
    label: string,
    selected: string,
    options: { key: string; text: string }[],
    onChange: (key: string) => void,
    disabled: boolean,
    className?: string
  ): React.ReactNode => {
    return (
      <AppleChoiceGroup
        label={label}
        selected={selected}
        options={options}
        onChange={onChange}
        disabled={disabled}
        className={className}
      />
    );
  };

  private _pulseYesNo = (target: HTMLElement): void => {
    const group = target.closest(`.${styles.segmented}`) as HTMLElement | null;
    if (!group || typeof group.animate !== 'function') {
      return;
    }
    group.animate(
      [
        { transform: 'scale(1)' },
        { transform: 'scale(0.96)', offset: 0.32 },
        { transform: 'scale(1)' }
      ],
      { duration: 280, easing: 'cubic-bezier(0.22, 1, 0.36, 1)' }
    );
  };

  private _persistHistory = (history: IFieldHistory): IFieldHistory => {
    saveFieldHistory(history);
    return history;
  };

  private _openHistory = (fieldId: string): void => {
    if (this._historyCloseTimer !== undefined) {
      window.clearTimeout(this._historyCloseTimer);
      this._historyCloseTimer = undefined;
    }
    this.setState({ historyFieldId: fieldId });
  };

  private _closeHistory = (): void => {
    this.setState({ historyFieldId: undefined });
  };

  private _onHistoryFieldBlur = (field: IFormField): void => {
    if (isHistoryTextField(field.label)) {
      const history = this._persistHistory(rememberFieldValue(this.state.history, field.label, field.value));
      this.setState({ history });
    }
    this._historyCloseTimer = window.setTimeout(() => {
      this._closeHistory();
    }, 160);
  };

  private _applyHistoryValue = (field: IFormField, value: string): void => {
    if (this._historyCloseTimer !== undefined) {
      window.clearTimeout(this._historyCloseTimer);
      this._historyCloseTimer = undefined;
    }
    this._onFieldValueChange(field.id, value);
    this.setState({ historyFieldId: undefined });
  };

  private _renderFieldHistory = (field: IFormField): React.ReactNode => {
    if (field.id !== this.state.historyFieldId || !isHistoryTextField(field.label)) {
      return undefined;
    }
    const suggestions = suggestionsFor(this.state.history, field.label, field.value);
    if (suggestions.length === 0) {
      return undefined;
    }
    return (
      <ul className={styles.historyList} role="listbox" aria-label={strings.RecentValuesLabel}>
        <li className={styles.historyCaption}>{strings.RecentValuesLabel}</li>
        {suggestions.map((value) => (
          <li key={value}>
            <button
              type="button"
              className={styles.historyItem}
              onMouseDown={(event) => {
                event.preventDefault();
                event.stopPropagation();
                this._applyHistoryValue(field, value);
              }}
            >
              {value}
            </button>
          </li>
        ))}
      </ul>
    );
  };

  private _renderBanner = (
    kind: 'error' | 'info' | 'success' | 'warning',
    message: React.ReactNode,
    onDismiss?: () => void
  ): React.ReactNode => {
    const kindClass = kind === 'error'
      ? styles.bannerError
      : kind === 'info'
        ? styles.bannerInfo
        : kind === 'success'
          ? styles.bannerSuccess
          : styles.bannerWarning;
    return (
      <div className={`${styles.banner} ${kindClass}`} role="status">
        <div className={styles.bannerText}>{message}</div>
        {onDismiss && (
          <button
            type="button"
            className={styles.bannerDismiss}
            onClick={onDismiss}
            aria-label={strings.Dismiss}
          >
            ×
          </button>
        )}
      </div>
    );
  };

  private _renderFieldName = (field: IFormField, documentKind: CorrespondenceKind): React.ReactNode => {
    const hint = this._fieldHint(field.label, documentKind);
    const text = isIssueDateField(field.label) ? ISSUE_DATE_DISPLAY_LABEL : field.label;
    return (
      <span className={styles.fieldName} title={hint}>
        {text}
        {isRequiredField(field.label) ? <span className={styles.required}> *</span> : undefined}
      </span>
    );
  };

  private _fieldHint = (label: string, documentKind: CorrespondenceKind): string | undefined => {
    if (isNameField(label)) {
      return documentKind === 'incoming'
        ? (strings.IncomingNameDescription || 'Generated when Incoming is selected. This field is read only.')
        : strings.NameDescription;
    }
    if (isRegistrationNumberField(label)) {
      return strings.RegistrationNumberDescription;
    }
    if (isProjectNumberField(label)) {
      return documentKind === 'incoming'
        ? (strings.IncomingProjectNumberDescription || 'From the text between Dear and Subject, matched to Notification Set-up Project Name.')
        : strings.ProjectNumberDescription;
    }
    if (isIssueDateField(label)) {
      return strings.IssueDateDescription || 'Date format: dd/MM/yyyy';
    }
    return undefined;
  };

  private _renderYesNo = (field: IFormField): React.ReactNode => {
    const selected = canonicalYesNo(field.value) || YES_VALUE;
    const isNo = selected === NO_VALUE;
    return (
      <div className={styles.yesNoRow}>
        <span className={styles.yesNoLabel}>
          {field.label}
          {isRequiredField(field.label) ? <span className={styles.required}> *</span> : undefined}
        </span>
        <div
          className={`${styles.segmented} ${isNo ? styles.segmentedNo : ''}`}
          role="radiogroup"
          aria-label={field.label}
        >
          <span className={styles.segmentThumb} aria-hidden={true} />
          {YES_NO_OPTIONS.map((name) => (
            <button
              key={name}
              type="button"
              className={`${styles.segment} ${selected === name ? styles.segmentActive : ''}`}
              aria-checked={selected === name}
              role="radio"
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                this._pulseYesNo(event.currentTarget);
                this._setActiveField(field.id);
                this._onFieldValueChange(field.id, selected === YES_VALUE ? NO_VALUE : YES_VALUE);
              }}
            >
              {name}
            </button>
          ))}
        </div>
      </div>
    );
  };

  private _requiredError = (field: IFormField, markRequired: boolean): string | undefined => {
    if (!this._isMissingRequired(field, markRequired)) {
      return undefined;
    }
    return strings.RequiredFieldError;
  };

  private _isMissingRequired = (field: IFormField, markRequired: boolean): boolean => {
    return markRequired && isRequiredField(field.label) && !(field.value || '').trim();
  };

  private _defaultActiveFieldId = (fields: IFormField[]): string | undefined => {
    const fillable = fields.filter((field) =>
      !isNameField(field.label) && !isRegistrationNumberField(field.label)
    )[0];
    if (fillable) {
      return fillable.id;
    }
    return fields.length > 0 ? fields[0].id : undefined;
  };

  private _defaultFieldValue = (label: string): string => {
    if (isSubProjectNumberField(label)) {
      return SUB_PROJECT_NONE;
    }
    if (isYesNoChoiceField(label)) {
      return YES_VALUE;
    }
    return '';
  };

  private _fieldsFromConfig = (config: string, existing?: IFormField[]): IFormField[] => {
    const labels = (config || '')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line.length > 0);
    const names = labels.length > 0 ? labels : DEFAULT_FORM_FIELDS;
    const valuesByLabel = new Map<string, string>();
    (existing || []).forEach((field) => {
      valuesByLabel.set(field.label, field.value);
    });
    return this._syncRegistrationFromName(names.map((label) => ({
      id: `field-${this._nextFieldId++}`,
      label,
      value: valuesByLabel.get(label) || this._defaultFieldValue(label)
    })));
  };

  private _setActiveField = (fieldId: string): void => {
    this.setState({ activeFieldId: fieldId, error: undefined });
  };

  private _isCalendarElement = (node: EventTarget | null): boolean => {
    if (!node || !(node instanceof Element)) {
      return false;
    }
    return !!node.closest('.ms-DatePicker-callout, .ms-Calendar, .ms-DatePicker');
  };

  private _patchCalendarButtons = (root?: ParentNode | null): void => {
    const scope = root || document;
    const buttons = scope.querySelectorAll('.ms-DatePicker-callout button, .ms-Calendar button');
    for (let i = 0; i < buttons.length; i++) {
      (buttons[i] as HTMLButtonElement).type = 'button';
    }
  };

  private _startCalendarObserver = (): void => {
    this._stopCalendarObserver();
    if (typeof MutationObserver === 'undefined') {
      return;
    }
    this._calendarObserver = new MutationObserver(() => {
      this._patchCalendarButtons();
    });
    this._calendarObserver.observe(document.body, { childList: true, subtree: true });
    this._patchCalendarButtons();
  };

  private _stopCalendarObserver = (): void => {
    if (this._calendarObserver) {
      this._calendarObserver.disconnect();
      this._calendarObserver = undefined;
    }
  };

  private _enablePreventScrollFocus = (): void => {
    if (this._originalFocus) {
      return;
    }
    this._originalFocus = HTMLElement.prototype.focus;
    const original = this._originalFocus;
    HTMLElement.prototype.focus = function (this: HTMLElement, options?: FocusOptions): void {
      original.call(this, { ...(options || {}), preventScroll: true });
    };
  };

  private _disablePreventScrollFocus = (): void => {
    if (this._focusPatchTimer !== undefined) {
      window.clearTimeout(this._focusPatchTimer);
      this._focusPatchTimer = undefined;
    }
    if (this._originalFocus) {
      HTMLElement.prototype.focus = this._originalFocus;
      this._originalFocus = undefined;
    }
  };

  private _onCalendarLayerMount = (): void => {
    this._calendarOpen = true;
    if (this._focusPatchTimer !== undefined) {
      window.clearTimeout(this._focusPatchTimer);
      this._focusPatchTimer = undefined;
    }
    this._enablePreventScrollFocus();
    this._startCalendarObserver();
  };

  private _onCalendarLayerUnmount = (): void => {
    this._calendarOpen = false;
    this._stopCalendarObserver();
    this._focusPatchTimer = window.setTimeout(() => {
      this._focusPatchTimer = undefined;
      if (!this._calendarOpen) {
        this._disablePreventScrollFocus();
      }
    }, 300);
  };

  private _onDocumentMouseDownCapture = (event: MouseEvent): void => {
    if (!this._isCalendarElement(event.target)) {
      return;
    }
    const button = (event.target as HTMLElement).closest('button');
    if (button) {
      (button as HTMLButtonElement).type = 'button';
    }
    this._patchCalendarButtons((event.target as HTMLElement).closest('.ms-DatePicker-callout, .ms-Calendar'));
  };

  private _onDocumentClickCapture = (event: MouseEvent): void => {
    if (!this._isCalendarElement(event.target)) {
      return;
    }
    const button = (event.target as HTMLElement).closest('button');
    if (button) {
      (button as HTMLButtonElement).type = 'button';
    }
  };

  private _onDocumentSubmitCapture = (event: Event): void => {
    if (!this._calendarOpen && !this._isCalendarElement(document.activeElement) && !this._isCalendarElement(event.target)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
  };

  private _onIssueDateSelect = (fieldId: string, date: Date | null | undefined): void => {
    this._enablePreventScrollFocus();
    this._onFieldValueChange(fieldId, date ? formatIssueDate(date) : '');
  };

  private _onFieldValueChange = (fieldId: string, value: string): void => {
    const target = this.state.fields.filter((field) => field.id === fieldId)[0];
    this.setState((prev) => {
      const current = prev.fields.filter((field) => field.id === fieldId)[0];
      if (current && isReadOnlyFormField(current.label)) {
        return {
          fields: prev.fields,
          activeFieldId: fieldId,
          showRequiredErrors: prev.showRequiredErrors,
          error: prev.error,
          uploadType: prev.uploadType
        };
      }
      const fields = prev.fields.map((field) => {
        if (field.id !== fieldId) {
          return field;
        }
        const nextValue = this._normalizeFieldValue(field.label, value);
        return {
          ...field,
          value: nextValue,
          debugSource: nextValue ? '手動輸入' : undefined
        };
      });
      const nextFields = this._syncRegistrationFromName(fields);
      const nextProjectNumber = this._namedValue(nextFields, isProjectNumberField);
      return {
        fields: nextFields,
        uploadType: canonicalUploadTypeForProjectNumber(prev.uploadType, nextProjectNumber),
        activeFieldId: fieldId,
        showRequiredErrors: false,
        error: undefined
      };
    }, () => {
      if (target && isProjectNumberField(target.label)) {
        this._refreshLeadingBlFromNotificationSetup(this._normalizeFieldValue(target.label, value));
      }
      if (target && isLeadingBlField(target.label) && this.state.incomingLetterType === 'email') {
        this._refreshEmailEoiProjectNumber(this._normalizeFieldValue(target.label, value));
      }
      if (target && (isProjectNumberField(target.label) || isLeadingBlField(target.label))) {
        this._refreshProjectUploadTypes();
      }
    });
  };

  private _refreshLeadingBlFromNotificationSetup = (projectNumber: string): void => {
    const file = this.state.file;
    if (!file) {
      return;
    }
    if (this.state.incomingLetterType === 'email' || isEoiProjectNumber(projectNumber)) {
      return;
    }
    const projectNo = sanitizeProjectNumber(projectNumber);
    if (!isValidProjectNumber(projectNo) || isEoiProjectNumber(projectNo)) {
      return;
    }
    if (this._leadingBlLookupTimer) {
      window.clearTimeout(this._leadingBlLookupTimer);
    }
    this._leadingBlLookupSeq = this._leadingBlLookupSeq + 1;
    const seq = this._leadingBlLookupSeq;
    this._leadingBlLookupTimer = window.setTimeout(() => {
      this._leadingBlLookupTimer = undefined;
      lookupLeadingBlFromNotificationSetup(this.props.spHttpClient, this.props.currentWebUrl, projectNo)
        .then((result) => {
          if (seq !== this._leadingBlLookupSeq) {
            return;
          }
          const warning = result.thresholdExceeded
            ? (strings.NotificationSetupThresholdHint ||
              'Could not read Leading BL from "Notification Set-up". Index the Project No column (this list has more than 5,000 items).')
            : undefined;
          if (!result.leadingBl && !result.projectNumber && !warning) {
            return;
          }
          this.setState((prev) => ({
            fields: this._applyNotificationSetupLookup(prev.fields, result),
            warning: warning || prev.warning
          }), () => {
            this._refreshProjectUploadTypes();
          });
        })
        .catch(() => {
          return;
        });
    }, 400);
  };

  private _applyNotificationSetupLookup = (fields: IFormField[], setup: INotificationSetupLookup): IFormField[] => {
    return fields.map((field) => {
      if (isProjectNumberField(field.label) && setup.projectNumber) {
        return {
          ...field,
          value: this._normalizeFieldValue(field.label, setup.projectNumber),
          debugSource: 'Notification Set-up'
        };
      }
      if (isLeadingBlField(field.label) && setup.leadingBl) {
        return {
          ...field,
          value: this._normalizeFieldValue(field.label, setup.leadingBl),
          debugSource: 'Notification Set-up'
        };
      }
      return field;
    });
  };

  private _refreshEmailEoiProjectNumber = (leadingBl: string): void => {
    if (this.state.incomingLetterType !== 'email') {
      return;
    }
    if (this._eoiLookupTimer) {
      window.clearTimeout(this._eoiLookupTimer);
    }
    this._eoiLookupSeq = this._eoiLookupSeq + 1;
    const seq = this._eoiLookupSeq;
    const bl = (leadingBl || '').trim();
    if (!bl) {
      this.setState((prev) => ({
        fields: this._setEmailEoiProjectNumber(prev.fields, ''),
        uploadType: canonicalUploadTypeForProjectNumber(prev.uploadType, '')
      }));
      return;
    }
    this._eoiLookupTimer = window.setTimeout(() => {
      this._eoiLookupTimer = undefined;
      this._lookupEmailEoiProjectNumber(bl).then((projectNumber) => {
        if (seq !== this._eoiLookupSeq) {
          return;
        }
        this.setState((prev) => ({
          fields: this._setEmailEoiProjectNumber(prev.fields, projectNumber),
          uploadType: canonicalUploadTypeForProjectNumber(prev.uploadType, projectNumber)
        }), () => {
          this._refreshProjectUploadTypes();
        });
      }).catch(() => {
        return;
      });
    }, 200);
  };

  private _lookupEmailEoiProjectNumber = async (leadingBl: string): Promise<string> => {
    const blSite = resolveLeadingBlSite(leadingBl, this.props.tenantUrl);
    const code = await lookupRootUrlMappingCode(this.props.spHttpClient, {
      listWebUrl: this.props.currentWebUrl,
      siteAbsoluteUrl: this.props.siteAbsoluteUrl,
      destinationSiteUrl: blSite ? blSite.siteUrl : '',
      leadingBl,
      listTitle: ROOT_URL_MAPPING_LIST_TITLE
    });
    return eoiProjectNumberFromCode(code);
  };

  private _setEmailEoiProjectNumber = (fields: IFormField[], projectNumber: string): IFormField[] => {
    return fields.map((field) => {
      if (!isProjectNumberField(field.label)) {
        return field;
      }
      const value = sanitizeProjectNumber(projectNumber);
      return {
        ...field,
        value,
        debugSource: value ? 'Root URL Mapping Code' : undefined
      };
    });
  };

  private _applyPdfFileName = (fields: IFormField[], fileName?: string): IFormField[] => {
    const pdfName = fileName ? nameFromPdfFile(fileName) : '';
    const withName = fields.map((field) => (
      isNameField(field.label)
        ? { ...field, value: pdfName, debugSource: pdfName ? 'PDF 檔名' : undefined }
        : field
    ));
    return this._syncRegistrationFromName(withName);
  };

  private _ensureIncomingName = (fields: IFormField[], regenerate?: boolean): IFormField[] => {
    const nameField = fields.filter((field) => isNameField(field.label))[0];
    const current = nameField ? (nameField.value || '').trim() : '';
    const value = regenerate || !current ? generateIncomingName() : current;
    const withName = fields.map((field) => (
      isNameField(field.label)
        ? { ...field, value, debugSource: 'Incoming 編號' }
        : field
    ));
    return this._syncRegistrationFromName(withName);
  };

  private _applyNameForKind = (fields: IFormField[], kind: CorrespondenceKind, fileName?: string, regenerateIncoming?: boolean): IFormField[] => {
    if (kind === 'incoming') {
      return this._ensureIncomingName(fields, regenerateIncoming);
    }
    return this._applyPdfFileName(fields, fileName);
  };

  private _fieldsForSelectedFile = (fields: IFormField[], fileName: string): IFormField[] => {
    const kind = correspondenceKindFromFileName(fileName);
    const next = kind === 'incoming'
      ? fields.map((field) => (
        isNameField(field.label) || isRegistrationNumberField(field.label)
          ? field
          : { ...field, value: this._defaultFieldValue(field.label), debugSource: undefined }
      ))
      : fields;
    const nameField = next.filter((field) => isNameField(field.label))[0];
    const regenerateIncoming = kind === 'incoming' && !isIncomingName(nameField ? nameField.value : '');
    return this._applyNameForKind(next, kind, fileName, regenerateIncoming);
  };

  private _syncRegistrationFromName = (fields: IFormField[]): IFormField[] => {
    const nameField = fields.filter((field) => isNameField(field.label))[0];
    const nameValue = nameField ? nameField.value : '';
    return fields.map((field) => (
      isRegistrationNumberField(field.label)
        ? { ...field, value: nameValue, debugSource: nameValue ? '從 Name 複製' : undefined }
        : field
    ));
  };

  private _normalizeFieldValue = (label: string, value: string): string => {
    value = stripOcrStyleTags(value).replace(/[ \t]+/g, ' ').trim();
    if (isLeadingBlField(label)) {
      return canonicalLeadingBl(value);
    }
    if (isProjectNumberField(label)) {
      return sanitizeProjectNumber(value);
    }
    if (isSubProjectNumberField(label)) {
      return canonicalSubProjectNumber(value);
    }
    if (isReceiverField(label)) {
      return this._stripHonorifics(this._stripParentheses(value));
    }
    if (isSenderField(label)) {
      const file = this.state.file;
      if (file && correspondenceKindFromFileName(file.name) === 'incoming') {
        return this._unwrapSenderIfFullyParenthesized(value);
      }
      return value;
    }
    if (isIssueDateField(label)) {
      return sanitizeIssueDate(value);
    }
    if (isYesNoChoiceField(label)) {
      return canonicalYesNo(value);
    }
    return value;
  };

  private _stripHonorifics = (value: string): string => {
    return (value || '')
      .replace(/^(?:(?:mr|mrs|ms|miss|dr|ir|prof(?:essor)?|engr?|sir|madam|mdm|mx|messrs)\b\.?\s*)+/i, '')
      .replace(/\s*(?:先生|女士|小姐|太太)\s*$/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  };

  private _unwrapSenderIfFullyParenthesized = (value: string): string => {
    const text = (value || '').trim();
    const wrapped = text.match(/^[(\uFF08]\s*([\s\S]+?)\s*[)\uFF09]$/);
    if (wrapped) {
      return (wrapped[1] || '').replace(/\s+/g, ' ').trim();
    }
    return text;
  };

  private _stripParentheses = (value: string): string => {
    let text = (value || '').trim();
    let previous = '';
    while (text !== previous) {
      previous = text;
      text = text
        .replace(/^[(\uFF08]\s*([\s\S]*?)\s*[)\uFF09]$/, '$1')
        .replace(/[(\uFF08][^)\uFF09]*[)\uFF09]/g, '')
        .replace(/[()\uFF08\uFF09]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
    return text;
  };

  private _fillActiveField = (text: string): void => {
    const trimmed = text.trim();
    if (!trimmed) {
      return;
    }
    const { activeFieldId, fields } = this.state;
    let targetId = activeFieldId || (fields.length > 0 ? fields[0].id : undefined);
    if (!targetId) {
      this.setState({ error: strings.SelectFieldFirst });
      return;
    }
    const target = fields.filter((field) => field.id === targetId)[0];
    if (target && isRegistrationNumberField(target.label)) {
      const nameField = fields.filter((field) => isNameField(field.label))[0];
      targetId = nameField ? nameField.id : targetId;
    }
    const fillTarget = fields.filter((field) => field.id === targetId)[0];
    if (fillTarget && isNameField(fillTarget.label)) {
      return;
    }
    const value = fillTarget ? this._normalizeFieldValue(fillTarget.label, trimmed) : trimmed;
    this.setState((prev) => {
      const fields = this._syncRegistrationFromName(prev.fields.map((field) => (
        field.id === targetId ? { ...field, value, debugSource: 'PDF 選取' } : field
      )));
      return {
        fields,
        fieldDebugMarks: buildOcrFieldMarks(prev.pages, fields),
        activeFieldId: targetId,
        error: undefined
      };
    });
  };

  private _uploadTypeLabel = (uploadType: UploadType): string => {
    if (uploadType === 'confidentialInvoice') {
      return strings.UploadTypeConfidentialInvoice || 'Confidential Invoice';
    }
    if (uploadType === 'confidentialMisc') {
      return strings.UploadTypeConfidentialMisc || 'Confidential MISC';
    }
    return strings.UploadTypeNormal || 'Normal';
  };

  private _onUploadTypeChange = (value: string): void => {
    const available = this.state.projectUploadTypes.length > 0
      ? this.state.projectUploadTypes
      : uploadTypeOptionsForProjectNumber(this._namedValue(this.state.fields, isProjectNumberField)).map((option) => option.key);
    this.setState({
      uploadType: constrainUploadType(value, available)
    });
  };

  private _refreshProjectUploadTypes = (fields?: IFormField[]): void => {
    const file = this.state.file;
    const list = fields || this.state.fields;
    const documentKind = file ? correspondenceKindFromFileName(file.name) : 'unknown';
    const projectDest = resolveUploadDestination(list, {
      tenantUrl: this.props.tenantUrl,
      libraryName: this.props.libraryName,
      folderPathTemplate: this.props.folderPathTemplate,
      uploadType: UPLOAD_TYPE_NORMAL,
      correspondenceKind: 'unknown'
    });
    const folderName = correspondenceFolderName(documentKind);
    if (!projectDest.siteUrl || !projectDest.folderPath || !folderName) {
      this.setState({ projectUploadTypes: [] });
      return;
    }
    if (this._uploadFolderLookupTimer) {
      window.clearTimeout(this._uploadFolderLookupTimer);
    }
    this._uploadFolderLookupSeq = this._uploadFolderLookupSeq + 1;
    const seq = this._uploadFolderLookupSeq;
    this._uploadFolderLookupTimer = window.setTimeout(() => {
      this._uploadFolderLookupTimer = undefined;
      lookupProjectUploadTypes(
        this.props.spHttpClient,
        projectDest.siteUrl,
        projectDest.libraryName,
        projectDest.folderPath,
        folderName
      ).then((types) => {
        if (seq !== this._uploadFolderLookupSeq) {
          return;
        }
        this.setState((prev) => ({
          projectUploadTypes: types,
          uploadType: constrainUploadType(prev.uploadType, types)
        }));
      }).catch(() => {
        if (seq !== this._uploadFolderLookupSeq) {
          return;
        }
        this.setState({ projectUploadTypes: [] });
      });
    }, 300);
  };

  private _labelTypeLabel = (labelType: LabelType): string => {
    if (labelType === 'confidential') {
      return strings.LabelTypeConfidential || 'Confidential';
    }
    if (labelType === 'invoice') {
      return strings.LabelTypeInvoice || 'Invoice';
    }
    if (labelType === 'site') {
      return strings.LabelTypeSite || 'Site';
    }
    return strings.LabelTypeNormal || 'Normal';
  };

  private _onLabelTypeChange = (value: string): void => {
    this.setState({
      labelType: canonicalLabelType(value)
    });
  };

  private _namedValue = (fields: IFormField[], match: (label: string) => boolean): string => {
    const field = fields.filter((item) => match(item.label))[0];
    return field ? (field.value || '').trim() : '';
  };

  private _issueDateIso = (value: string): string => {
    const date = parseIssueDate(value);
    if (!date) {
      return '';
    }
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${date.getFullYear()}-${month < 10 ? '0' : ''}${month}-${day < 10 ? '0' : ''}${day}`;
  };

  private _subProjectNumberOptions = (): IDropdownOption[] => {
    return SUB_PROJECT_NUMBER_OPTIONS.map((name) => ({
      key: name,
      text: name
    }));
  };

  private _leadingBlOptions = (currentValue?: string): IDropdownOption[] => {
    const options: IDropdownOption[] = LEADING_BL_OPTIONS.map((name) => ({
      key: name,
      text: name
    }));
    const current = (currentValue || '').trim();
    if (current && LEADING_BL_OPTIONS.indexOf(current) < 0) {
      options.unshift({
        key: current,
        text: `${current} (not a mapped site)`
      });
    }
    return options;
  };

  private _onPdfSelectText = (text: string, indexes: number[]): void => {
    this.setState({ selectedWordIndexes: indexes });
    this._fillActiveField(text);
  };

  private _openFilePicker = (): void => {
    this._fileInput.current?.click();
  };

  private _onFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const selected = event.target.files && event.target.files.length > 0
      ? event.target.files[0]
      : undefined;

    if (!selected) {
      return;
    }

    if (!selected.name.toLowerCase().endsWith('.pdf') && selected.type !== 'application/pdf') {
      this._originalPdfBytes = undefined;
      this.setState({
        error: strings.InvalidFileType,
        file: undefined
      });
      event.target.value = '';
      return;
    }

    this._keepOriginalPdf(selected).catch(() => {
      // Errors are surfaced in component state.
    });
  };

  private _keepOriginalPdf = async (selected: File): Promise<void> => {
    this._revokePageUrls(this.state.pages);
    try {
      this._originalPdfBytes = new Uint8Array(await selected.arrayBuffer());
    } catch {
      this._originalPdfBytes = undefined;
      this.setState({ error: strings.InvalidFileType, file: undefined });
      return;
    }

    this.setState({
      file: selected,
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      successFolderUrl: undefined,
      warning: undefined,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      fields: this._fieldsForSelectedFile(this.state.fields, selected.name),
      showRequiredErrors: false,
      incomingLetterType: '',
      dearToSubjectText: '',
      projectUploadTypes: []
    }, () => {
      this._refreshProjectUploadTypes();
    });
  };

  private _onToggleOcrStyles = (): void => {
    this.setState((prev) => ({
      showOcrStyles: !prev.showOcrStyles
    }));
  };

  private _onRestylePage = (): void => {
    this._restyleCurrentPage().catch(() => {
      // Errors are surfaced in component state.
    });
  };

  private _restyleCurrentPage = async (): Promise<void> => {
    const page = this.state.pages[Math.max(0, this.state.currentPage - 1)];
    const source = this._originalPdfBytes
      ? this._originalPdfBytes.slice()
      : this.state.file;
    if (!page || !source) {
      this.setState({
        error: strings.RestyleNeedConvert || 'Convert first, then restyle this page.'
      });
      return;
    }
    const seq = this._restyleSeq + 1;
    this._restyleSeq = seq;
    this.setState({
      isRestyling: true,
      error: undefined,
      info: undefined
    });
    try {
      const styled = await PdfOcrService.restylePage(source, page.pageNumber, page.words || []);
      if (seq !== this._restyleSeq) {
        return;
      }
      this.setState((prev) => ({
        pages: prev.pages.map((item) => (
          item.pageNumber === page.pageNumber
            ? { ...item, words: styled.words, styleSpans: styled.styleSpans }
            : item
        )),
        showOcrStyles: true,
        isRestyling: false,
        info: locFormat(
          strings.RestylePageDone || 'Restyled page {0}. Blue = underline, red = separator.',
          'Restyled page {0}. Blue = underline, red = separator.',
          String(page.pageNumber)
        )
      }));
    } catch (err) {
      if (seq !== this._restyleSeq) {
        return;
      }
      const message = err instanceof Error ? err.message : '';
      this.setState({
        isRestyling: false,
        error: message && message !== 'Error' ? message : (strings.RestyleFailed || 'Could not restyle this page.')
      });
    }
  };

  private _onDevToolsOpenChange = (open: boolean): void => {
    this.setState({
      devToolsOpen: open
    });
  };

  private _onPreviousPage = (): void => {
    this.setState((prev) => ({
      currentPage: Math.max(1, prev.currentPage - 1),
      selectedWordIndexes: []
    }));
  };

  private _onNextPage = (): void => {
    this.setState((prev) => ({
      currentPage: Math.min(prev.pages.length, prev.currentPage + 1),
      selectedWordIndexes: []
    }));
  };

  private _onConvert = (): void => {
    this._runOcr().catch(() => {
      // Errors are surfaced in component state.
    });
  };

  private _onUpload = (): void => {
    this._uploadToSharePoint().catch(() => {
      // Errors are surfaced in component state.
    });
  };

  private _uploadToSharePoint = async (): Promise<void> => {
    const { file, fields } = this.state;
    if (!file) {
      this.setState({ error: strings.UploadNeedFile });
      return;
    }

    const requiredMissing = missingRequiredFields(fields);
    if (requiredMissing.length > 0) {
      this.setState({
        showRequiredErrors: true,
        error: locFormat(strings.UploadMissingFields, 'Fill these required fields before uploading: {0}', requiredMissing.join(', '))
      });
      return;
    }

    const projectNumber = this._namedValue(fields, isProjectNumberField);
    const available = this.state.projectUploadTypes.length > 0
      ? this.state.projectUploadTypes
      : uploadTypeOptionsForProjectNumber(projectNumber).map((option) => option.key);
    const resolvedUploadType = constrainUploadType(this.state.uploadType, available);
    const destination = resolveUploadDestination(fields, {
      tenantUrl: this.props.tenantUrl,
      libraryName: this.props.libraryName,
      folderPathTemplate: this.props.folderPathTemplate,
      uploadType: resolvedUploadType,
      correspondenceKind: correspondenceKindFromFileName(file.name)
    });

    if (destination.missingFields.length > 0) {
      this.setState({
        error: locFormat(strings.UploadMissingFields, 'Fill these required fields before uploading: {0}', destination.missingFields.join(', '))
      });
      return;
    }

    if (destination.unrecognizedLeadingBl) {
      this.setState({
        error: locFormat(strings.UploadUnknownLeadingBl, 'Leading BL "{0}" does not match a business line site. Choose a listed business line.', destination.unrecognizedLeadingBl)
      });
      return;
    }

    const projectNumberField = fields.filter((field) => isProjectNumberField(field.label))[0];
    if (projectNumberField && !isValidProjectNumber(projectNumberField.value)) {
      this.setState({ error: strings.UploadInvalidProjectNumber });
      return;
    }

    this.setState({
      isUploading: true,
      error: undefined,
      success: undefined,
      successUrl: undefined,
      successFolderUrl: undefined,
      warning: undefined,
      showRequiredErrors: false,
      uploadStatus: strings.UploadStarting
    });

    try {
      let pdfBytes = this._originalPdfBytes && this._originalPdfBytes.byteLength > 0
        ? this._originalPdfBytes
        : new Uint8Array(await file.arrayBuffer());
      let labelWarning = '';
      const isIncoming = correspondenceKindFromFileName(file.name) === 'incoming';
      if (isIncoming) {
        try {
          this.setState({ uploadStatus: strings.UploadGeneratingLabel || 'Generating label page…' });
          const projectNo = this._namedValue(fields, isProjectNumberField);
          const staff = isEoiProjectNumber(projectNo)
            ? emptyLabelStaff()
            : await lookupLabelStaffFromNotificationSetup(
              this.props.spHttpClient,
              this.props.currentWebUrl,
              projectNo
            );
          const labelPng = await generateLabelPagePng(this.props.spHttpClient, {
            labelType: this.state.labelType,
            projectNumber: this._namedValue(fields, isProjectNumberField),
            leadingBl: this._namedValue(fields, isLeadingBlField),
            registrationNumber: this._namedValue(fields, isRegistrationNumberField) || this._namedValue(fields, isNameField),
            organization: this._namedValue(fields, isOrganizationField),
            sender: this._namedValue(fields, isSenderField),
            receiver: this._namedValue(fields, isReceiverField),
            subject: this._namedValue(fields, isSubjectField),
            subProjectNumber: this._namedValue(fields, isSubProjectNumberField) || SUB_PROJECT_NONE,
            issueDateIso: this._issueDateIso(this._namedValue(fields, isIssueDateField)),
            refNo: this._namedValue(fields, isRefNoField),
            hasAttachment: canonicalYesNo(this._namedValue(fields, isAttachmentField)) === YES_VALUE,
            ccToAecom: canonicalYesNo(this._namedValue(fields, isCcToAecomField)) === YES_VALUE,
            uploadType: resolvedUploadType,
            staff,
            hasScan: canonicalYesNo(this._namedValue(fields, isScanField)) === YES_VALUE,
            siteUrls: [this.props.siteAbsoluteUrl, this.props.currentWebUrl].filter((url, index, list) =>
              !!url && list.indexOf(url) === index
            )
          });
          pdfBytes = await appendLabelPageToPdf(pdfBytes, await labelPng.arrayBuffer());
        } catch (labelErr) {
          const details = labelErr instanceof Error ? labelErr.message : '';
          labelWarning = details
            ? `${strings.UploadLabelFailed || 'Label page could not be added; the original PDF was uploaded.'} ${details}`
            : (strings.UploadLabelFailed || 'Label page could not be added; the original PDF was uploaded.');
        }
      }

      const service = new SharePointUploadService(this.props.spHttpClient);
      const result = await service.uploadPdf(
        destination,
        file,
        fileNameFromFields(file, fields),
        (status) => {
          this.setState({ uploadStatus: this._localizeUploadStatus(status) });
        },
        this.props.currentWebUrl,
        fields.map((field) => ({ label: field.label, value: field.value })),
        pdfBytes
      );
      const scanNote = isIncoming && canonicalYesNo(this._namedValue(fields, isScanField)) === YES_VALUE
        ? (strings.UploadBlankPageAdded || 'A blank label page was added at the end of the PDF.')
        : '';
      this.setState({
        isUploading: false,
        uploadStatus: undefined,
        success: locFormat(strings.UploadSucceeded, 'Uploaded {0}.', result.fileName) + (scanNote ? `\n${scanNote}` : ''),
        successUrl: result.fileUrl,
        successFolderUrl: result.folderUrl,
        warning: [labelWarning, result.metadataError].filter((item) => !!item).join('\n') || undefined,
        history: this._persistHistory(rememberFieldValues(this.state.history, fields))
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : strings.UploadFailed;
      this.setState({
        isUploading: false,
        uploadStatus: undefined,
        error: message
      });
    }
  };

  private _runOcr = async (): Promise<void> => {
    const { file } = this.state;
    if (!file) {
      return;
    }

    this._revokePageUrls(this.state.pages);
    this.setState({
      isProcessing: true,
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      successFolderUrl: undefined,
      warning: undefined,
      showRequiredErrors: false,
      incomingLetterType: '',
      dearToSubjectText: '',
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      progress: {
        page: 0,
        totalPages: 0,
        percent: 0,
        status: strings.LoadingEngine
      }
    });

    try {
      const documentKind = correspondenceKindFromFileName(file.name);
      const result = await PdfOcrService.extractText(
        this._originalPdfBytes ? this._originalPdfBytes.slice() : file,
        'eng',
        (progress) => {
          this.setState({
            progress: {
              ...progress,
              status: this._localizeStatus(progress.status)
            }
          });
        },
        (page) => {
          this.setState((prev) => ({
            pages: prev.pages.concat([page]),
            currentPage: page.pageNumber
          }));
        },
        documentKind === 'incoming'
          ? createIncomingOcrDecider()
          : undefined
      );

      const previewPages = documentKind === 'incoming'
        ? keepIncomingEmailPreviewPages(result.pages)
        : result.pages;
      let filled: {
        fields: IFormField[];
        info: string | undefined;
        warning?: string;
        letterType: string;
        dearToSubjectText: string;
      };
      this.setState({
        progress: {
          page: previewPages.length,
          totalPages: previewPages.length,
          percent: 100,
          status: strings.ExtractingFields
        }
      });
      try {
        filled = await this._fillFields(result.pages, documentKind);
      } catch {
        filled = { fields: this.state.fields, info: undefined, letterType: '', dearToSubjectText: '' };
      }
      this.setState({
        pages: previewPages,
        currentPage: 1,
        isProcessing: false,
        progress: undefined,
        fields: filled.fields,
        fieldDebugMarks: buildOcrFieldMarks(previewPages, filled.fields),
        error: undefined,
        info: filled.info,
        warning: filled.warning,
        incomingLetterType: documentKind === 'incoming' ? (filled.letterType || '') : '',
        dearToSubjectText: filled.dearToSubjectText || '',
        uploadType: canonicalUploadTypeForProjectNumber(
          this.state.uploadType,
          this._namedValue(filled.fields, isProjectNumberField)
        )
      }, () => {
        this._refreshProjectUploadTypes(filled.fields);
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : '';
      this.setState({
        isProcessing: false,
        progress: undefined,
        error: message && message !== 'Error' ? message : strings.OcrFailed
      });
    }
  };

  private _fillFields = async (pages: IOcrPageResult[], kind: CorrespondenceKind): Promise<{
    fields: IFormField[];
    error: string | undefined;
    info: string | undefined;
    warning?: string;
    letterType: string;
    dearToSubjectText: string;
  }> => {
    const labels = this.state.fields.map((field) => field.label);
    const incoming = kind === 'incoming';
    const detected = incoming
      ? await detectIncomingFields(pages || [])
      : await detectOutgoingFields(pages || []);
    let keywordValues: { [label: string]: string } = {};
    try {
      keywordValues = extractFieldValues(pages || [], labels);
    } catch {
      keywordValues = {};
    }
    let aiValues: { [label: string]: string } = {};
    let info: string | undefined;
    const aiConfig = {
      endpoint: this.props.azureOpenAiEndpoint || '',
      apiKey: this.props.azureOpenAiApiKey || '',
      deployment: this.props.azureOpenAiDeployment || '',
      apiVersion: this.props.azureOpenAiApiVersion || '2024-08-01-preview'
    };

    if (isAiExtractionConfigured(aiConfig)) {
      try {
        const ocrText = (pages || []).map((page) => formatOcrTextWithStyles(page.words || []) || page.text || '').join('\n');
        aiValues = await extractFieldsWithAi(
          ocrText,
          labels,
          aiConfig,
          incoming ? incomingAiHints(detected) : outgoingAiHints(detected)
        );
      } catch {
        aiValues = {};
      }
    }

    let fields = this.state.fields;
    try {
      const mapped = this.state.fields.map((field) => {
        if (isNameField(field.label)) {
          return field;
        }
        const pickedField = incoming
          ? pickIncomingFieldValue(field.label, detected, aiValues[field.label] || '', keywordValues[field.label] || '')
          : pickOutgoingFieldValue(field.label, detected, aiValues[field.label] || '', keywordValues[field.label] || '');
        return {
          ...field,
          value: this._normalizeFieldValue(field.label, pickedField.value),
          debugSource: pickedField.source || undefined
        };
      });
      const mappedName = mapped.filter((field) => isNameField(field.label))[0];
      fields = incoming
        ? this._ensureIncomingName(mapped, !isIncomingName(mappedName ? mappedName.value : ''))
        : this._applyPdfFileName(mapped, this.state.file ? this.state.file.name : undefined);
    } catch {
      fields = this.state.fields;
    }

    let warning: string | undefined;
    const incomingEmail = incoming && detected.letterType === 'email';
    if (incoming && !incomingEmail) {
      const projectNameHint = (detected.projectNameHint || detected.agreementNo || '').trim();
      if (projectNameHint) {
        try {
          const setup = await lookupNotificationSetupByProjectName(
            this.props.spHttpClient,
            this.props.currentWebUrl,
            projectNameHint
          );
          fields = this._applyNotificationSetupLookup(fields, setup);
          if (setup.projectNumber && !setup.leadingBl) {
            const leading = await lookupLeadingBlFromNotificationSetup(
              this.props.spHttpClient,
              this.props.currentWebUrl,
              setup.projectNumber
            );
            fields = this._applyNotificationSetupLookup(fields, leading);
            if (leading.thresholdExceeded) {
              warning = strings.NotificationSetupThresholdHint ||
                'Could not read Leading BL from "Notification Set-up". Index the Project No column (this list has more than 5,000 items).';
            }
          }
          if (setup.thresholdExceeded) {
            warning = strings.NotificationSetupThresholdHint ||
              'Could not read Leading BL from "Notification Set-up". Index the Project No column (this list has more than 5,000 items).';
          }
        } catch {
          warning = undefined;
        }
      }
    } else if (incomingEmail) {
      const leadingField = fields.filter((field) => isLeadingBlField(field.label))[0];
      const leadingBl = leadingField ? (leadingField.value || '').trim() : '';
      if (leadingBl) {
        try {
          const projectNumber = await this._lookupEmailEoiProjectNumber(leadingBl);
          fields = this._setEmailEoiProjectNumber(fields, projectNumber);
        } catch {
          fields = this._setEmailEoiProjectNumber(fields, '');
        }
      } else {
        fields = this._setEmailEoiProjectNumber(fields, '');
      }
    } else {
      const projectField = fields.filter((field) => isProjectNumberField(field.label))[0];
      const resolvedProjectNumber = projectField ? sanitizeProjectNumber(projectField.value) : '';
      if (resolvedProjectNumber) {
        try {
          const setup = await lookupLeadingBlFromNotificationSetup(
            this.props.spHttpClient,
            this.props.currentWebUrl,
            resolvedProjectNumber
          );
          if (setup.leadingBl || setup.projectNumber) {
            fields = this._applyNotificationSetupLookup(fields, setup);
          }
          if (setup.thresholdExceeded) {
            warning = strings.NotificationSetupThresholdHint ||
              'Could not read Leading BL from "Notification Set-up". Index the Project No column (this list has more than 5,000 items).';
          }
        } catch {
          warning = undefined;
        }
      }
    }

    return {
      error: undefined,
      info,
      warning,
      fields,
      letterType: incoming ? (detected.letterType || '') : '',
      dearToSubjectText: incoming && detected.letterType !== 'email'
        ? (detected.projectNameHint || '').trim()
        : ''
    };
  };

  private _localizeStatus = (status: string): string => {
    if (status.indexOf('Loading OCR engine') === 0) {
      return strings.LoadingEngine;
    }
    if (status.indexOf('Rendering page') === 0) {
      return status.replace('Rendering page', strings.RenderingPage).replace(' of ', ` ${strings.Of} `);
    }
    if (status.indexOf('OCR page') === 0) {
      return status.replace('OCR page', strings.OcrPage).replace(' of ', ` ${strings.Of} `);
    }
    if (status === 'Completed') {
      return strings.Completed;
    }
    if (status === strings.ExtractingFields) {
      return strings.ExtractingFields;
    }
    return status;
  };

  private _localizeUploadStatus = (status: string): string => {
    if (status.indexOf('Generating label') === 0) {
      return strings.UploadGeneratingLabel || 'Generating label page…';
    }
    if (status.indexOf('Checking Root URL Mapping') === 0) {
      return strings.UploadCheckingMapping;
    }
    if (status.indexOf('Reading destination') === 0) {
      return strings.UploadReadingSite;
    }
    if (status.indexOf('Finding Project Documents') === 0) {
      return strings.UploadFindingLibrary;
    }
    if (status.indexOf('Creating folder') === 0) {
      return strings.UploadCreatingFolder;
    }
    if (status.indexOf('Uploading PDF') === 0) {
      return strings.UploadSendingFile;
    }
    if (status.indexOf('Updating library fields') === 0) {
      return strings.UploadUpdatingFields;
    }
    return status;
  };

  private _onClear = (): void => {
    if (this._fileInput.current) {
      this._fileInput.current.value = '';
    }
    this._originalPdfBytes = undefined;
    this._restyleSeq = this._restyleSeq + 1;
    this._revokePageUrls(this.state.pages);
    this.setState({
      file: undefined,
      pages: [],
      currentPage: 1,
      selectedWordIndexes: [],
      fields: this.state.fields.map((field) => ({
        ...field,
        value: this._defaultFieldValue(field.label),
        debugSource: undefined
      })),
      fieldDebugMarks: [],
      error: undefined,
      info: undefined,
      success: undefined,
      successUrl: undefined,
      successFolderUrl: undefined,
      warning: undefined,
      isUploading: false,
      uploadStatus: undefined,
      showRequiredErrors: false,
      isRestyling: false,
      progress: undefined,
      uploadType: UPLOAD_TYPE_NORMAL,
      labelType: LABEL_TYPE_NORMAL,
      incomingLetterType: '',
      dearToSubjectText: '',
      projectUploadTypes: []
    });
  };

  private _revokePageUrls = (pages: IOcrPageResult[]): void => {
    (pages || []).forEach((page) => {
      const url = page && page.imageUrl ? page.imageUrl : '';
      if (url.indexOf('blob:') === 0) {
        URL.revokeObjectURL(url);
      }
    });
  };

  private _openInNewTab = (event: React.MouseEvent<HTMLElement>, url: string): void => {
    if (!url) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    window.open(url, '_blank', 'noopener,noreferrer');
  };

  private _clearError = (): void => {
    this.setState({ error: undefined });
  };

  private _clearInfo = (): void => {
    this.setState({ info: undefined });
  };

  private _clearSuccess = (): void => {
    this.setState({ success: undefined, successUrl: undefined, successFolderUrl: undefined });
  };

  private _clearWarning = (): void => {
    this.setState({ warning: undefined });
  };
}
