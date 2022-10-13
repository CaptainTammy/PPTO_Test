/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global document, Office */

Office.onReady((info) => {
  write("loaded");
  if (info.host === Office.HostType.PowerPoint) {
    write("loaded powerpoint");
    document.getElementById("run").onclick = run;
  }
});

// Function that writes to a div with id='message' on the page.
function write(message) {
  console.log('************------------' + message + '-----------*******************');
  document.getElementById("message").innerText += message;
}

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  write('Adding a shape');

  await PowerPoint.run(async (context) => {
    write('Load slides');
    context.presentation.load('slides');
    await context.sync();
    write('Slides loaded ' + context.presentation.slides);
    write('Getting shapes');
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    write('Shapes ' + shapes);
    await context.sync();
    write('Shapes loaded');

    const braces = shapes.addGeometricShape(PowerPoint.GeometricShapeType.bracePair);
    await context.sync();
    braces.left = 100;
    braces.top = 400;
    braces.height = 50;
    braces.width = 150;
    braces.name = "Braces";
    braces.fill.setSolidColor("lightblue");
    braces.textFrame.textRange.text = "Shape text";
    braces.textFrame.textRange.font.color = "purple";
    braces.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;
    write('Shape added');
    return context.sync();
  });
}
