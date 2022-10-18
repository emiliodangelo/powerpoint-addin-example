import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Office, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {
    return PowerPoint.run(async (context) => {
      try {
        const presentation = context.presentation;

        // Add new slide
        presentation.slides.add();
        await context.sync();

        // If we wait a couple seconds here we avoid an error in PowerPoint Online
        // await sleep(2000);

        // Get the number of slides in presentation
        const slideCount = presentation.slides.getCount();
        await context.sync();

        // Get a reference to the last slide
        // Need to disable validation because the object doesn't have a load method
        // eslint-disable-next-line office-addins/load-object-before-read
        const slide = presentation.slides.getItemAt(slideCount.value - 1);
        slide.shapes.load("items/$none");
        await context.sync();

        // Delete all shapes in the slide
        slide.shapes.items.forEach((shape) => {
          shape.delete();
        });
        await context.sync();

        // Add title to slide
        const titleOptions: PowerPoint.ShapeAddOptions = {
          top: 25,
          left: 25,
          height: 50,
          width: 915,
        };
        const title = slide.shapes.addTextBox("Lorem ipsum dolor sit amet", titleOptions);
        title.textFrame.textRange.font.name = "Calibri Light";
        title.textFrame.textRange.font.size = 28;
        title.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeTextToFitShape;

        // Add description to slide
        const descriptionOptions: PowerPoint.ShapeAddOptions = {
          top: 80,
          left: 330,
          height: 415,
          width: 610,
        };
        const description = slide.shapes.addTextBox("Nam tristique feugiat enim, vitae faucibus libero fermentum quis. Curabitur hendrerit laoreet justo, eget congue justo feugiat imperdiet. Interdum et malesuada fames ac ante ipsum primis in faucibus. Aliquam erat volutpat. Vestibulum vel ex gravida, interdum nisl eu, interdum eros. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus blandit feugiat est, pharetra molestie turpis. Praesent nisl leo, blandit at odio ac, luctus commodo libero. Aenean a augue nunc. Nulla facilisi. In hac habitasse platea dictumst. Proin eget odio volutpat, maximus sem ut, laoreet nisi. Phasellus condimentum ex lacus, nec efficitur nisl tempor vestibulum.", descriptionOptions);
        description.textFrame.textRange.font.name = "Calibri Light";
        description.textFrame.textRange.font.size = 16;
        description.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeTextToFitShape;

        // Go to the last slide
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Slide, async () => {
          const pictureOptions = {
            coercionType: Office.CoercionType.Image,
            imageTop: 80,
            imageLeft: 25,
            imageWidth: 300,
          };

          // Add a picture in the selected slide
          Office.context.document.setSelectedDataAsync(
            "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9Im5vIj8+CjwhRE9DVFlQRSBzdmcgUFVCTElDICItLy9XM0MvL0RURCBTVkcgMS4xLy9FTiIgImh0dHA6Ly93d3cudzMub3JnL0dyYXBoaWNzL1NWRy8xLjEvRFREL3N2ZzExLmR0ZCI+Cjxzdmcgd2lkdGg9IjEwMCUiIGhlaWdodD0iMTAwJSIgdmlld0JveD0iMCAwIDMyIDMyIiB2ZXJzaW9uPSIxLjEiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIHhtbDpzcGFjZT0icHJlc2VydmUiIHhtbG5zOnNlcmlmPSJodHRwOi8vd3d3LnNlcmlmLmNvbS8iIHN0eWxlPSJmaWxsLXJ1bGU6ZXZlbm9kZDtjbGlwLXJ1bGU6ZXZlbm9kZDtzdHJva2UtbGluZWpvaW46cm91bmQ7c3Ryb2tlLW1pdGVybGltaXQ6MjsiPgogICAgPHBhdGggZD0iTTMwLDMuNDE0TDI4LjU4NiwyTDIsMjguNTg2TDMuNDE0LDMwTDUuNDE0LDI4TDI2LDI4QzI3LjA5NywyNy45OTkgMjcuOTk5LDI3LjA5NyAyOCwyNkwyOCw1LjQxNEwzMCwzLjQxNFpNMjYsMjZMNy40MTQsMjZMMTUuMjA3LDE4LjIwN0wxNy41ODYsMjAuNTg2QzE4LjM2MiwyMS4zNjEgMTkuNjM4LDIxLjM2MSAyMC40MTQsMjAuNTg2TDIyLDE5TDI2LDIyLjk5N0wyNiwyNlpNMjYsMjAuMTY4TDIzLjQxNCwxNy41ODJDMjIuNjM4LDE2LjgwNyAyMS4zNjIsMTYuODA3IDIwLjU4NiwxNy41ODJMMTksMTkuMTY4TDE2LjYyMywxNi43OTFMMjYsNy40MTRMMjYsMjAuMTY4WiIgc3R5bGU9ImZpbGw6cmdiKDE5MiwxOTIsMTkyKTtmaWxsLXJ1bGU6bm9uemVybzsiLz4KICAgIDxwYXRoIGQ9Ik02LDIyTDYsMTlMMTEsMTQuMDAzTDEyLjM3MywxNS4zNzdMMTMuNzg5LDEzLjk2MUwxMi40MTQsMTIuNTg2QzExLjYzOCwxMS44MSAxMC4zNjIsMTEuODEgOS41ODYsMTIuNTg2TDYsMTYuMTcyTDYsNkwyMiw2TDIyLDRMNiw0QzQuOTAzLDQuMDAxIDQuMDAxLDQuOTAzIDQsNkw0LDIyTDYsMjJaIiBzdHlsZT0iZmlsbDpyZ2IoMTkyLDE5MiwxOTIpO2ZpbGwtcnVsZTpub256ZXJvOyIvPgogICAgPHJlY3QgaWQ9Il9UcmFuc3BhcmVudF9SZWN0YW5nbGVfIiB4PSIwIiB5PSIwIiB3aWR0aD0iMzIiIGhlaWdodD0iMzIiIHN0eWxlPSJmaWxsOm5vbmU7Ii8+Cjwvc3ZnPgo=",
            pictureOptions
          );
        });
      } catch (error) {
        console.error(error);
      }
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Run
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}
