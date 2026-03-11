async (page) => {
  const appUrl = "http://127.0.0.1:8501";
  const pdfPath = "/Users/griffin/Downloads/Domain Money Mock Financial Plan.pdf";

  await page.goto(appUrl, { waitUntil: "domcontentloaded", timeout: 120000 });
  await page.getByText("Upload PDF").waitFor({ timeout: 120000 });

  const fileInput = page.locator('input[type="file"]').first();
  await fileInput.setInputFiles(pdfPath);
  await page.getByRole("button", { name: "Build workbook" }).click();

  const outcome = await Promise.race([
    page
      .getByRole("button", { name: "Download review report" })
      .waitFor({ timeout: 900000 })
      .then(() => "report-ready"),
    page
      .getByText("Run failed", { exact: false })
      .waitFor({ timeout: 900000 })
      .then(() => "run-failed"),
  ]);

  if (outcome === "run-failed") {
    const bodyText = (await page.locator("body").textContent()) || "";
    throw new Error(`HollyPlanner surfaced a failure state.\n\n${bodyText}`);
  }

  const workbookButton = page.getByRole("button", { name: "Download filled workbook" });
  const reviewButton = page.getByRole("button", { name: "Download review report" });
  const bodyText = (await page.locator("body").textContent()) || "";

  return {
    reviewReportReady: await reviewButton.isVisible(),
    workbookReady: await workbookButton.isVisible().catch(() => false),
    headline: bodyText.includes("Workbook ready")
      ? "Workbook ready"
      : bodyText.includes("Review required")
        ? "Review required"
        : "unknown",
  };
}
