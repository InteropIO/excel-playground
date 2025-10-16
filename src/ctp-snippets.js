ctpDescriptor = {
  id: "validationScrollBox",
  ui: {
    type: "ScrollBox",
    id: "validationScrollBox",
    backColor: "#2E2E2E", // dark gray
    children: Array.from({ length: 25 }, (_, i) => ({
      type: "Label",
      id: `lbl${i + 1}`,
      text: `A${i + 1} should be a positive int`,
      foreColor: "#FF0000", // red
    })),
  },
  dockPosition: "msoCTPDockPositionLeft",
};

xlService.createOrUpdateCTP(null, ctpDescriptor, console.info);
