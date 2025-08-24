/* global Word */

import * as React from "react";
import { makeStyles, tokens, Button } from "@fluentui/react-components";
import { NORD } from "../utils/seed/nord";

interface AppProps {}

const useStyles = makeStyles({
  root: {
    height: "100vh",
  },
  container: {
    display: "flex",
    minHeight: "100%",
    flex: 1,
  },
  sidebar: {
    display: "grid",
    borderRight: `1px solid ${tokens.colorPaletteRedBorder1}`,
    width: "200px",
    minHeight: "100%",
  },
  colorGrid: {
    display: "grid",
    height: "fit-content",
    gap: "8px",
  },
  colorRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  colorName: {
    width: "80px",
  },
  colorSwatch: {
    height: "36px",
    width: "80px",
    position: "relative",
    cursor: "pointer",
    ":hover": {
      "& .hover-overlay": {
        opacity: 1,
      },
    },
  },
  hoverOverlay: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(0, 0, 0, 0.7)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "4px",
    opacity: 0,
    transition: "opacity 0.2s ease",
  },
  hoverButton: {
    minWidth: "24px",
    height: "20px",
    fontSize: "10px",
    padding: "2px 4px",
  },
  shadeLabel: {
    color: tokens.colorNeutralForeground1,
  },
});

const App: React.FC<AppProps> = () => {
  const styles = useStyles();

  // Function to apply highlight color to selected text
  const applyHighlightColor = async (hex: string) => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.font.highlightColor = hex;
        debugger;
        await context.sync();
      });
    } catch (error) {
      console.error("Error applying highlight color:", error);
    }
  };

  // Function to apply text color to selected text
  const applyTextColor = async (hex: string) => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.font.color = hex;
        await context.sync();
      });
    } catch (error) {
      console.error("Error applying text color:", error);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.container}>
        <div className={styles.sidebar}>
          <h1>Hello World XZZXZ</h1>
        </div>
        <div className={styles.colorGrid}>
          {NORD.map(({ swatches, name }) => (
            <div key={name} className={styles.colorRow}>
              <span className={styles.colorName}>{name}</span>
              {swatches.map(({ shade, hex }) => (
                <div
                  id={`color-tile-${shade}`}
                  key={shade}
                  className={styles.colorSwatch}
                  style={{ backgroundColor: hex }}
                >
                  <div className={styles.shadeLabel}>{shade}</div>
                  <div className={`${styles.hoverOverlay} hover-overlay`}>
                    <Button
                      className={styles.hoverButton}
                      size="small"
                      onClick={() => applyHighlightColor(hex)}
                      title={`Apply ${shade} as highlight color`}
                    >
                      H
                    </Button>
                    <Button
                      className={styles.hoverButton}
                      size="small"
                      onClick={() => applyTextColor(hex)}
                      title={`Apply ${shade} as text color`}
                    >
                      T
                    </Button>
                  </div>
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default App;
