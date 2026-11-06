import {
  CheckCircle2Icon,
  FileArchiveIcon,
  FileCode2Icon,
  FileSearchIcon,
  FolderOpenIcon,
  HammerIcon,
  ImageIcon,
  PresentationIcon,
  SearchIcon,
  TriangleAlertIcon,
  XCircleIcon,
} from "lucide-react";
import { useRef, useState } from "react";

import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Field,
  FieldContent,
  FieldDescription,
  FieldGroup,
  FieldLabel,
  FieldLegend,
  FieldSet,
  FieldTitle,
} from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Spinner } from "@/components/ui/spinner";
import { Textarea } from "@/components/ui/textarea";
import type { Builder } from "@/lib/useBuilder";
import { cn } from "@/lib/utils";

function isOutputFormat(value: unknown): value is "pptx" | "png" {
  return value === "pptx" || value === "png";
}

function StepHeading({
  step,
  title,
  description,
}: {
  step: number;
  title: string;
  description: string;
}) {
  return (
    <div className="flex items-start gap-2">
      <Badge variant="secondary" className="size-5 rounded-full p-0">
        {step}
      </Badge>
      <div className="flex flex-col gap-0.5">
        <FieldTitle>{title}</FieldTitle>
        <FieldDescription>{description}</FieldDescription>
      </div>
    </div>
  );
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function ResourceStatus({ status }: { status: Builder["resourceStatus"] }) {
  if (status.kind === "idle") return null;

  if (status.kind === "reading") {
    return (
      <Alert aria-live="polite">
        <Spinner />
        <AlertTitle>Reading resources</AlertTitle>
        <AlertDescription>{status.message}</AlertDescription>
      </Alert>
    );
  }

  const Icon = status.kind === "success" ? CheckCircle2Icon : XCircleIcon;

  return (
    <Alert variant={status.kind === "error" ? "destructive" : "success"} aria-live="polite">
      <Icon />
      <AlertTitle>
        {status.kind === "success" ? "Resources ready" : "Resources could not be read"}
      </AlertTitle>
      <AlertDescription>{status.message}</AlertDescription>
    </Alert>
  );
}

export function BuildPanel({ builder }: { builder: Builder }) {
  const zipInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  return (
    <Card>
      <CardHeader>
        <CardTitle>Build your presentation</CardTitle>
        <CardDescription>
          Work through the six steps, then forge an editable PowerPoint.
        </CardDescription>
      </CardHeader>
      <CardContent>
        <FieldGroup>
          <Field>
            <StepHeading step={1} title="Presentation name" description="Name your presentation" />
            <Input
              id="presentation-name"
              aria-label="Presentation name"
              placeholder="e.g. Lesson 6 — Understanding Scams"
              value={builder.presentationName}
              onChange={(event) => builder.setPresentationName(event.target.value)}
            />
          </Field>

          <Field>
            <StepHeading
              step={2}
              title="Resources"
              description="Select the exported resources folder or ZIP linked from the lesson HTML"
            />
            <Button
              type="button"
              variant="outline"
              className={cn(
                "h-auto w-full flex-col border-dashed py-6 text-muted-foreground",
                dragOver && "border-primary bg-primary/5 text-foreground",
              )}
              onClick={() => zipInputRef.current?.click()}
              onDragOver={(event) => {
                event.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(event) => {
                event.preventDefault();
                setDragOver(false);
                const file = event.dataTransfer.files[0];
                if (file) void builder.handleZipFile(file);
              }}
            >
              <FileArchiveIcon />
              <span>
                <strong className="font-medium text-foreground">Choose a file</strong> or drag and
                drop a resources ZIP
              </span>
            </Button>
            <input
              ref={zipInputRef}
              type="file"
              accept=".zip,application/zip"
              className="sr-only"
              tabIndex={-1}
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) void builder.handleZipFile(file);
              }}
            />
            <div className="flex items-center gap-3 text-xs text-muted-foreground">
              <span className="h-px flex-1 bg-border" />
              or
              <span className="h-px flex-1 bg-border" />
            </div>
            <Button
              type="button"
              variant="outline"
              className="h-auto w-full flex-col whitespace-normal py-4 text-center text-muted-foreground"
              onClick={() => folderInputRef.current?.click()}
            >
              <FolderOpenIcon />
              <span>
                Choose the exported resources folder —
                <strong className="ml-1 font-medium text-foreground">no zipping needed</strong>
              </span>
            </Button>
            <input
              ref={folderInputRef}
              type="file"
              webkitdirectory=""
              directory=""
              multiple
              className="sr-only"
              tabIndex={-1}
              onChange={(event) => {
                const files = event.target.files;
                if (files?.length) void builder.handleFolderFiles(files);
                event.target.value = "";
              }}
            />
            <ResourceStatus status={builder.resourceStatus} />

            {builder.lastZipInfo && (
              <Alert>
                <FileArchiveIcon />
                <AlertTitle>Use your last resources file?</AlertTitle>
                <AlertDescription className="flex flex-col gap-3">
                  <p>
                    {builder.lastZipInfo.name} ({formatBytes(builder.lastZipInfo.size)}) is stored
                    in this browser.
                  </p>
                  <div className="flex flex-wrap gap-2">
                    <Button size="sm" onClick={() => void builder.selectLastZip()}>
                      Use this file
                    </Button>
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => {
                        builder.dismissLastZipSuggestion();
                        zipInputRef.current?.click();
                      }}
                    >
                      Choose a different file
                    </Button>
                  </div>
                </AlertDescription>
              </Alert>
            )}
          </Field>

          <Field>
            <StepHeading
              step={3}
              title="Paste lesson HTML"
              description="Paste the lesson HTML you copied from Seqta"
            />
            <Textarea
              id="lesson-html"
              aria-label="Lesson HTML"
              placeholder="Paste the raw lesson HTML here…"
              className="min-h-32 font-mono text-xs"
              value={builder.lessonHTML}
              onChange={(event) => builder.setLessonHTML(event.target.value)}
            />
            <Button
              variant="outline"
              size="sm"
              className="self-start"
              disabled={!builder.lessonHTML.trim()}
              onClick={builder.openPreview}
            >
              <SearchIcon data-icon="inline-start" />
              Preview parsed content
            </Button>
          </Field>

          <Field>
            <StepHeading
              step={4}
              title="Extra files (optional)"
              description="For a lesson with no resource list you can paste — search your loaded resources by the lesson's own words, then tick the ones that belong"
            />
            <Button
              type="button"
              variant="outline"
              size="sm"
              className="self-start"
              disabled={!builder.resources}
              onClick={builder.findExtraFiles}
            >
              <FileSearchIcon data-icon="inline-start" />
              Find matching files
            </Button>
            {!builder.resources && (
              <FieldDescription>Load a resources folder or ZIP in step 2 first.</FieldDescription>
            )}

            {builder.hasSearchedExtraFiles && builder.resourceCandidates.length === 0 && (
              <Alert aria-live="polite">
                <TriangleAlertIcon />
                <AlertTitle>No likely matches found</AlertTitle>
                <AlertDescription>
                  Try adding a bit more of the lesson's text or title above, then search again.
                </AlertDescription>
              </Alert>
            )}

            {builder.resourceCandidates.length > 0 && (
              <div className="flex flex-col gap-2 rounded-md border p-3">
                {builder.resourceCandidates.map((candidate) => {
                  const path = candidate.entry.relativePath;
                  return (
                    <FieldLabel key={path} className="flex-row items-start gap-2">
                      <Checkbox
                        checked={builder.selectedExtraFiles.has(path)}
                        onCheckedChange={() => builder.toggleExtraFile(path)}
                      />
                      <span className="flex min-w-0 flex-1 flex-col text-sm">
                        <span className="truncate font-medium text-foreground">
                          {path.split("/").pop()}
                        </span>
                        <span className="text-xs text-muted-foreground">
                          {formatBytes(candidate.entry.bytes)}
                        </span>
                      </span>
                    </FieldLabel>
                  );
                })}
              </div>
            )}
          </Field>

          <FieldSet>
            <FieldLegend className="flex items-start gap-2">
              <Badge variant="secondary" className="size-5 rounded-full p-0">
                5
              </Badge>
              <span>Choose output</span>
            </FieldLegend>
            <FieldDescription>PowerPoint or lesson pages</FieldDescription>
            <RadioGroup
              value={builder.outputFormat}
              onValueChange={(value) => {
                if (isOutputFormat(value)) builder.setOutputFormat(value);
              }}
              className="grid-cols-1 sm:grid-cols-2"
            >
              <FieldLabel>
                <Field orientation="horizontal">
                  <RadioGroupItem value="pptx" />
                  <PresentationIcon />
                  <FieldContent>
                    <FieldTitle>PowerPoint</FieldTitle>
                    <FieldDescription>Editable .pptx presentation</FieldDescription>
                  </FieldContent>
                </Field>
              </FieldLabel>
              <FieldLabel data-disabled="true">
                <Field orientation="horizontal" data-disabled="true">
                  <RadioGroupItem value="png" disabled />
                  <ImageIcon />
                  <FieldContent>
                    <FieldTitle className="flex items-center gap-2">
                      Lesson pages <Badge variant="outline">Coming soon</Badge>
                    </FieldTitle>
                    <FieldDescription>Individual .png images</FieldDescription>
                  </FieldContent>
                </Field>
              </FieldLabel>
            </RadioGroup>
          </FieldSet>

          <Field>
            <StepHeading
              step={6}
              title="Forge presentation"
              description="Build and download the presentation"
            />
            <Button
              className="w-full"
              disabled={builder.forgeState.kind === "forging"}
              onClick={builder.forge}
            >
              {builder.forgeState.kind === "forging" ? (
                <Spinner data-icon="inline-start" />
              ) : (
                <HammerIcon data-icon="inline-start" />
              )}
              {builder.forgeState.kind === "forging" ? "Forging…" : "Forge presentation"}
            </Button>

            {builder.forgeState.kind === "forging" && (
              <Alert aria-live="polite">
                <FileCode2Icon />
                <AlertTitle>Building your presentation</AlertTitle>
                <AlertDescription>{builder.forgeState.message}</AlertDescription>
              </Alert>
            )}

            {builder.forgeState.kind === "success" && (
              <Alert variant="success" aria-live="polite">
                <CheckCircle2Icon />
                <AlertTitle>Your presentation is ready</AlertTitle>
                <AlertDescription>
                  It has been downloaded to your browser. Open the PowerPoint and review it before
                  sharing.
                </AlertDescription>
              </Alert>
            )}

            {builder.forgeState.kind === "error" && (
              <Alert variant="destructive" aria-live="assertive">
                <TriangleAlertIcon />
                <AlertTitle>Presentation could not be built</AlertTitle>
                <AlertDescription>{builder.forgeState.message}</AlertDescription>
              </Alert>
            )}
          </Field>
        </FieldGroup>
      </CardContent>
    </Card>
  );
}
