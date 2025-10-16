xlService.write({ range: "A1" }, "Expected: Single string value in A2");
xlService.write({ range: "A2" }, "Hello");

xlService.write(
  { range: "A4" },
  "Expected: Single numeric value 1 filled into A5:C7"
);
xlService.write({ range: "A5:C7" }, 1);

xlService.write(
  { range: "A8" },
  "Expected: Flat horizontal array [1,2,3] -> A9:C9"
);
xlService.write({ range: "A9:C9" }, [1, 2, 3]);

xlService.write(
  { range: "A11" },
  "Expected: Flat vertical array [4,5,6] -> A12:A14"
);
xlService.write({ range: "A12:A14" }, [4, 5, 6]);

xlService.write(
  { range: "A16" },
  "Expected: Rectangular matrix 2x2 -> A17:B18"
);
xlService.write({ range: "A17:B18" }, [
  [1, 2],
  [3, 4],
]);

xlService.write({ range: "A20" }, "Expected: Jagged array (padded) -> A21:C23");
xlService.write({ range: "A21:C23" }, [[1], [2, 3], [4, 5, 6]]);

xlService.write({ range: "A25" }, "Expected: Matrix of strings -> A26:B27");
xlService.write({ range: "A26:B27" }, [
  ["a", "b"],
  ["c", "d"],
]);

xlService.write(
  { range: "A29" },
  "Expected: Vertical array of strings -> A30:A32"
);
xlService.write({ range: "A30:A32" }, ["x", "y", "z"]);

xlService.write(
  { range: "A34" },
  "Expected: Horizontal array of strings -> A35:C35"
);
xlService.write({ range: "A35:C35" }, ["p", "q", "r"]);

// horizontal resize
xlService.write(
  { range: "A37", resizeOrientation: "horizontal" },
  Array.from({ length: 50 }, (_, i) => [`Row ${i + 1}`, i])
);

// vertical resize
xlService.write(
  { range: "A40", resizeOrientation: "vertical" },
  Array.from({ length: 5 }, (_, i) => [`Row ${i + 1}`, i])
);

// vertical resize
xlService.write(
  { range: "A1", resizeOrientation: "vertical" },
  Array.from({ length: 50 * 1000 }, (_, i) => [`Row ${i + 1}`, i])
);
// horizontal resize
xlService.write(
  { range: "A1", resizeOrientation: "horizontal" },
  Array.from({ length: 50 }, (_, i) => [`Row ${i + 1}`, i])
);
