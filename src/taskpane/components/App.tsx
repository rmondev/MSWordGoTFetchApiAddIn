import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
//import TextInsertion from "./TextInsertion";
import { Button, Field, makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertBooksIntoDocument, insertAliasesIntoDocument } from "../taskpane";
//import TextInsertion from "./TextInsertion";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
  },
  button: {
    backgroundColor: "#2b579a", // Fluent UI Primary Color
    color: "black",
    fontSize: "16px",
    padding: "10px 20px",
    borderRadius: "5px",
    marginTop: "20px",
    ":hover": {
      backgroundColor: "#005A9E", // Darker shade on hover
    },
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Hello" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      {/* Press Button to Generate Book Titles */}
      <Field>Hello</Field>
      {/* Button centred on the page */}
      <Button appearance="primary" onClick={insertBooksIntoDocument}>
        Insert Books
      </Button>
      <Button appearance="outline" onClick={insertAliasesIntoDocument}>
        Insert Aliases
      </Button>
    </div>
  );
};

export default App;
