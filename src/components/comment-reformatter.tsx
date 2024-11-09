'use client'

import { useState } from 'react'
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle, CardDescription, CardFooter } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { FileUp, Download } from 'lucide-react'

export function CommentReformatterComponent() {
  const [file, setFile] = useState<File | null>(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [processedFileUrl, setProcessedFileUrl] = useState<string | null>(null)
  const [status, setStatus] = useState('')

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files[0]) {
      setFile(event.target.files[0])
      setStatus('File uploaded successfully.')
      setProcessedFileUrl(null)
    }
  }

  const handleReformat = async () => {
    if (!file) {
      setStatus('Please upload a file first.')
      return
    }

    setIsProcessing(true)
    setStatus('Processing file...')

    try {
      // Create FormData object to send file
      const formData = new FormData()
      formData.append('file', file)

      // Send file to server
      const response = await fetch('/api/reformat-comments', {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        throw new Error('Failed to process file')
      }

      // Get the processed file from the response
      const processedFile = await response.blob()
      setProcessedFileUrl(URL.createObjectURL(processedFile))
      setStatus('File processed successfully.')
    } catch (error) {
      setStatus('Error processing file: ' + (error as Error).message)
    } finally {
      setIsProcessing(false)
    }
  }

  return (
    <Card className="w-full max-w-md mx-auto">
      <CardHeader>
        <CardTitle>FDOT Comment Reformatter</CardTitle>
        <CardDescription>Upload and reformat FDOT ERC ThreadReports</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="space-y-2">
          <Label htmlFor="file-upload">Upload ThreadReport</Label>
          <div className="relative w-full h-20 border-2 border-dashed border-gray-300 rounded-lg">
            <Input
              id="file-upload"
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            <div className="flex items-center justify-center h-full">
              <span className="text-gray-500">{file ? file.name : "Choose a ThreadReport or drag it here"}</span>
            </div>
          </div>
        </div>
        <Button onClick={handleReformat} disabled={!file || isProcessing} className="w-full">
          {isProcessing ? 'Processing...' : 'Reformat Spreadsheet'}
          <FileUp className="w-4 h-4 ml-2" />
        </Button>
        {status && <p className="text-sm text-center">{status}</p>}
      </CardContent>
      <CardFooter>
        {processedFileUrl && (
          <Button asChild className="w-full">
            <a href={processedFileUrl} download="reformatted_comments.xlsx">
              Download Reformatted Spreadsheet
              <Download className="w-4 h-4 ml-2" />
            </a>
          </Button>
        )}
      </CardFooter>
    </Card>
  )
}