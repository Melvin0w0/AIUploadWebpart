import * as React from 'react';
import { IOcrPageResult } from '../services/IPdfOcr';
import { joinOcrWords, wordIndexAtPoint, wordsInRect } from '../services/ocrSelection';
import styles from './AiUpload.module.scss';

export interface IPdfHighlightViewerProps {
  page: IOcrPageResult;
  selectedIndexes: number[];
  onSelectText: (text: string, indexes: number[]) => void;
}

interface IPdfHighlightViewerState {
  displayWidth: number;
  displayHeight: number;
  isDragging: boolean;
  startX: number;
  startY: number;
  currentX: number;
  currentY: number;
}

export default class PdfHighlightViewer extends React.Component<IPdfHighlightViewerProps, IPdfHighlightViewerState> {
  private _imageRef: React.RefObject<HTMLImageElement>;

  public constructor(props: IPdfHighlightViewerProps) {
    super(props);
    this._imageRef = React.createRef<HTMLImageElement>();
    this.state = {
      displayWidth: 0,
      displayHeight: 0,
      isDragging: false,
      startX: 0,
      startY: 0,
      currentX: 0,
      currentY: 0
    };
  }

  public componentDidMount(): void {
    window.addEventListener('resize', this._syncSize);
    window.addEventListener('mouseup', this._onWindowMouseUp);
  }

  public componentWillUnmount(): void {
    window.removeEventListener('resize', this._syncSize);
    window.removeEventListener('mouseup', this._onWindowMouseUp);
  }

  public componentDidUpdate(prevProps: IPdfHighlightViewerProps): void {
    if (prevProps.page.imageUrl !== this.props.page.imageUrl) {
      this._syncSize();
    }
  }

  public render(): React.ReactElement<IPdfHighlightViewerProps> {
    const { page, selectedIndexes } = this.props;
    const { displayWidth, displayHeight, isDragging, startX, startY, currentX, currentY } = this.state;
    const words = page.words || [];
    const scale = page.width > 0 && displayWidth > 0 ? displayWidth / page.width : 1;
    const dragX = Math.min(startX, currentX) * scale;
    const dragY = Math.min(startY, currentY) * scale;
    const dragWidth = Math.abs(currentX - startX) * scale;
    const dragHeight = Math.abs(currentY - startY) * scale;

    return (
      <div className={styles.pdfViewer}>
        <div
          className={styles.pageStage}
          onMouseDown={this._onMouseDown}
          onMouseMove={this._onMouseMove}
        >
          <img
            ref={this._imageRef}
            className={styles.pageImage}
            src={page.imageUrl}
            alt=""
            draggable={false}
            onLoad={this._syncSize}
          />
          {displayWidth > 0 && (
            <svg
              className={styles.highlightOverlay}
              width={displayWidth}
              height={displayHeight}
            >
              {words.map((word, index) => {
                const isSelected = selectedIndexes.indexOf(index) >= 0;
                return (
                  <rect
                    key={`${word.x0}-${word.y0}-${index}`}
                    x={word.x0 * scale}
                    y={word.y0 * scale}
                    width={Math.max(1, (word.x1 - word.x0) * scale)}
                    height={Math.max(1, (word.y1 - word.y0) * scale)}
                    className={isSelected ? styles.wordSelected : styles.wordHit}
                  />
                );
              })}
              {isDragging && dragWidth > 2 && dragHeight > 2 && (
                <rect
                  x={dragX}
                  y={dragY}
                  width={dragWidth}
                  height={dragHeight}
                  className={styles.dragRect}
                />
              )}
            </svg>
          )}
        </div>
      </div>
    );
  }

  private _syncSize = (): void => {
    const image = this._imageRef.current;
    if (!image) {
      return;
    }
    this.setState({
      displayWidth: image.clientWidth,
      displayHeight: image.clientHeight
    });
  };

  private _toImagePoint = (event: React.MouseEvent<HTMLDivElement> | MouseEvent): { x: number; y: number } | undefined => {
    const image = this._imageRef.current;
    const { page } = this.props;
    if (!image || page.width === 0 || image.clientWidth === 0) {
      return undefined;
    }
    const bounds = image.getBoundingClientRect();
    const scaleX = page.width / bounds.width;
    const scaleY = page.height / bounds.height;
    return {
      x: (event.clientX - bounds.left) * scaleX,
      y: (event.clientY - bounds.top) * scaleY
    };
  };

  private _onMouseDown = (event: React.MouseEvent<HTMLDivElement>): void => {
    if (event.button !== 0) {
      return;
    }
    const point = this._toImagePoint(event);
    if (!point) {
      return;
    }
    event.preventDefault();
    this.setState({
      isDragging: true,
      startX: point.x,
      startY: point.y,
      currentX: point.x,
      currentY: point.y
    });
  };

  private _onMouseMove = (event: React.MouseEvent<HTMLDivElement>): void => {
    if (!this.state.isDragging) {
      return;
    }
    const point = this._toImagePoint(event);
    if (!point) {
      return;
    }
    this.setState({
      currentX: point.x,
      currentY: point.y
    });
  };

  private _onWindowMouseUp = (event: MouseEvent): void => {
    if (!this.state.isDragging) {
      return;
    }
    const point = this._toImagePoint(event) || {
      x: this.state.currentX,
      y: this.state.currentY
    };
    this._finishSelection(point.x, point.y);
  };

  private _finishSelection = (endX: number, endY: number): void => {
    const { page, onSelectText } = this.props;
    const { startX, startY } = this.state;
    const words = page.words || [];
    const movement = Math.max(Math.abs(endX - startX), Math.abs(endY - startY));

    let indexes: number[];
    if (movement < 4) {
      const wordIndex = wordIndexAtPoint(words, endX, endY);
      indexes = wordIndex >= 0 ? [wordIndex] : [];
    } else {
      indexes = wordsInRect(words, {
        x0: startX,
        y0: startY,
        x1: endX,
        y1: endY
      });
    }

    this.setState({
      isDragging: false,
      currentX: endX,
      currentY: endY
    });

    if (indexes.length > 0) {
      const selectedWords = indexes.map((index) => words[index]).filter((word) => !!word);
      onSelectText(joinOcrWords(selectedWords), indexes);
    }
  };
}
