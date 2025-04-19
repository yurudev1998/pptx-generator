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
    const base64 = buffer.toString("base64");

    return res.status(200).json({ filename: "Weekly_Report.pptx", base64 });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Failed to generate PowerPoint" });
  }
}
