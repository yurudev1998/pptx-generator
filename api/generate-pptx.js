import pptxgen from "pptxgenjs";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Only POST requests allowed");
  }

  try {
    const { tasks } = req.body;

    const pptx = new pptxgen();

    const slide = pptx.addSlide();
    slide.addText("Weekly Task Report", { x: 0.5, y: 0.3, fontSize: 24, bold: true });

    const tableData = [
      ["Task", "Status", "Assignee", "Description"],
      ...tasks.map(task => [
        task.Task,
        task.Status,
        task.Assignee,
        task.Description || "No Description"
      ])
    ];

    slide.addTable(tableData, {
      x: 0.5,
      y: 1,
      w: 9,
      colW: [3, 1.5, 2, 3],
      border: { pt: "1", color: "666666" },
      fill: "F1F1F1"
    });

    const buffer = await pptx.write("nodebuffer");

    // Set headers for file download
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", 'attachment; filename="Weekly_Report.pptx"');
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to generate PowerPoint" });
  }
}
