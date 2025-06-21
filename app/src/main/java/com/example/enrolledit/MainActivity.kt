package com.example.enrolledit

import android.app.Activity
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.provider.OpenableColumns
import android.util.Log
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat.checkSelfPermission
import androidx.lifecycle.lifecycleScope
import com.airbnb.lottie.LottieAnimationView
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException



@Suppress("DEPRECATION")
class MainActivity : AppCompatActivity() {

    private val PICK_FILE_REQUEST_CODE = 1
    private var docxFileUri: Uri? = null
    private val REQUEST_CODE_PERMISSIONS = 1001

    // Debugging function
   private fun logDebug(tag: String, message: String) {
        Log.d(tag, message)
    }


    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        checkPermissions()

        val selectFileButton: Button = findViewById(R.id.selectFileButton)
        val generateFilesButton: Button = findViewById(R.id.generateFilesButton)
        val startNumberEditText: EditText = findViewById(R.id.startNumberEditText)
        val endNumberEditText: EditText = findViewById(R.id.endNumberEditText)
        val prefixEditText: EditText = findViewById(R.id.prefixEditText)




        selectFileButton.setOnClickListener {
            // Launch file picker
            val intent = Intent(Intent.ACTION_OPEN_DOCUMENT).apply {
                addCategory(Intent.CATEGORY_OPENABLE)
                type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
            startActivityForResult(intent, PICK_FILE_REQUEST_CODE)
        }
        generateFilesButton.setOnClickListener {
            val prefix = prefixEditText.text.toString().trim()
            val startNumber = startNumberEditText.text.toString()
            val endNumber = endNumberEditText.text.toString()
            if (docxFileUri != null && startNumber.isNotEmpty() && endNumber.isNotEmpty()) {
                lifecycleScope.launch {
                    showLoadingAnimation(true)
                    //    generateFiles(docxFileUri!!, startNumber, endNumber)
                    generateFiles(docxFileUri!!, prefix, startNumber, endNumber)
                    showLoadingAnimation(false)
                }
            } else {
                Toast.makeText(
                    this,
                    "Please select a file and enter both start and end enrollment numbers",
                    Toast.LENGTH_SHORT
                ).show()
            }
        }
    }

    private fun showLoadingAnimation(show: Boolean) {
        val loadingAnimationView = findViewById<LottieAnimationView>(R.id.loadingAnimationView)
        if (show) {
            loadingAnimationView.visibility = View.VISIBLE
            loadingAnimationView.playAnimation()
        } else {

            loadingAnimationView.visibility = View.GONE
            loadingAnimationView.cancelAnimation()
        }
    }

    private val requiredPermissions = arrayOf(
        android.Manifest.permission.WRITE_EXTERNAL_STORAGE,
        android.Manifest.permission.READ_EXTERNAL_STORAGE
    )

    private fun checkPermissions() {
        if (checkSelfPermission(
                this,
                android.Manifest.permission.WRITE_EXTERNAL_STORAGE
            ) != PackageManager.PERMISSION_GRANTED
        ) {
            ActivityCompat.requestPermissions(this, requiredPermissions, REQUEST_CODE_PERMISSIONS)
        }
    }

    @Deprecated("Deprecated in Java")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if (requestCode == PICK_FILE_REQUEST_CODE && resultCode == Activity.RESULT_OK) {
            data?.data?.let {
                docxFileUri = it
                val fileName = getFileName(it)
                val selectedFileTextView: TextView = findViewById(R.id.filename)
                selectedFileTextView.visibility = View.VISIBLE
                selectedFileTextView.text = fileName
            }
        }
    }

    private fun getFileName(it: Uri): String {
        var fileName = ""
        val cursor = contentResolver.query(it, null, null, null, null)
        cursor?.use {
            if (it.moveToFirst()) {
                val nameIndex = it.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                if (nameIndex != -1) {
                    fileName = it.getString(nameIndex)
                    Toast.makeText(this, "File Selected Successfully", Toast.LENGTH_SHORT).show()
                }
            }
        }
        return fileName
    }

    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<out String>,
        grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == REQUEST_CODE_PERMISSIONS) {
            if (grantResults.isNotEmpty() && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                // Permission granted
            } else {
                // Permission denied
                Toast.makeText(this, "Permission denied to access storage", Toast.LENGTH_SHORT)
                    .show()
            }
        }
    }


   //THIS IS REGEX CODE WORKING
   private suspend fun generateFiles(
        uri: Uri,
        prefix: String,
        startNumber: String,
        endNumber: String
    ) =
        withContext(Dispatchers.IO) {
            try {
                val inputStream = contentResolver.openInputStream(uri) ?: return@withContext
                val tempDocxFile = File.createTempFile("sample", ".docx", cacheDir)
                logDebug("temp",tempDocxFile.absolutePath)
                val outputStream = FileOutputStream(tempDocxFile)
                inputStream.copyTo(outputStream)
                outputStream.close()
                inputStream.close()

                val startNum = startNumber.toInt()
                val endNum = endNumber.toInt()

                // Create a new directory to store the generated files
                val externalStorageDir =
                    File(
                        Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS),
                        "EnrollEdits"
                    )
                if (!externalStorageDir.exists()) {
                    if (!externalStorageDir.mkdirs()) {
                        throw IOException("Failed to create directory")
                    }
                }
                Log.d(
                    "MainActivity",
                    "Generated files will be stored in ${externalStorageDir.absolutePath}"
                )
                val enrollList = listOf(Regex("""\d{2}BE[A-Z]{2}\d{5}"""),Regex("""\d{2}[A-Z]{5}\d{5}"""))

                for (num in startNum..endNum) {
                    val newDocxFile = File(externalStorageDir, "$prefix$num.docx")
                    val newHeader = "$prefix$num"

                    FileInputStream(tempDocxFile).use { fis ->
                        val document = XWPFDocument(fis)
                        // Replace headers in header sections
                        for (header in document.headerList) {
                            for (paragraph in header.paragraphs) {
                                for (run in paragraph.runs) {
                                    val text = run.text()
                                    logDebug("Header", "Original header text: $text")
                                    // Match the enrollment number pattern and replace it
                                    for(patterns in enrollList) {
                                        val updatedText = text.replace(
                                            patterns,
                                            newHeader
                                        )
                                        logDebug("Header", "Updated header text: $updatedText")
                                        if (updatedText != text) {
                                            run.setText(updatedText, 0)
                                            logDebug("Header", "Header text replaced")
                                        } else {
                                            logDebug("Header", "No match found for replacement")
                                        }
                                    }
                                }
                            }
                        }
                        FileOutputStream(newDocxFile).use { fos ->
                            document.write(fos)
                        }
                        document.close()
                    }
                    Log.d("MainActivity", "File generated: ${newDocxFile.absolutePath}")
                }
                withContext(Dispatchers.Main) {
                    Toast.makeText(
                        this@MainActivity,
                        "Files generated successfully",
                        Toast.LENGTH_SHORT
                    ).show()
                }
            } catch (e: IOException) {
                e.printStackTrace()
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Error: ${e.message}", Toast.LENGTH_SHORT)
                        .show()
                }
            }
        }

}

/*  This is Orignal code for 22BECE

      private suspend fun generateFiles(uri: Uri, startNumber: String, endNumber: String) =
          withContext(Dispatchers.IO) {
              try {
                  val inputStream = contentResolver.openInputStream(uri) ?: return@withContext
                  val tempDocxFile = File.createTempFile("sample", ".docx", cacheDir)
                  val outputStream = FileOutputStream(tempDocxFile)
                  inputStream.copyTo(outputStream)
                  outputStream.close()
                  inputStream.close()

                  val startNum = startNumber.toInt()
                  val endNum = endNumber.toInt()

                  // Create a new directory to store the generated files
                  val externalStorageDir =
                      File(
                          Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS),
                          "EnrollEdits"
                      )
                  if (!externalStorageDir.exists()) {
                      if (!externalStorageDir.mkdirs()) {
                          throw IOException("Failed to create directory")
                      }
                  }

                  Log.d(
                      "MainActivity",
                      "Generated files will be stored in: ${externalStorageDir.absolutePath}"
                  )

                  for (num in startNum..endNum) {
                      val newDocxFile = File(externalStorageDir, "22BECE30$num.docx")
                      val newHeader = "22BECE30$num"

                      FileInputStream(tempDocxFile).use { fis ->
                          val document = XWPFDocument(fis)
                          for (header in document.headerList) {
                              for (paragraph in header.paragraphs) {
                                  paragraph.runs.forEach { run ->
                                      val text = run.text()
                                      Log.d("Header", "Original header text: $text")
                                      if (text.contains("22BECE")) {
                                          run.setText(text.replace(Regex("\\d+"), newHeader), 0)
                                          Log.d("Header", "Updated header text: ${run.text()}")

                                      }
                                  }
                              }
                          }
                          // Update header on all pages
                          for (paragraph in document.paragraphs) {
                              paragraph.runs.forEach { run ->
                                  val text = run.text()
                                  if (text.contains("22BECE")) {
                                      run.setText(text.replace(Regex("22BECE\\d+"), newHeader), 0)
                                  }
                              }
                          }
                          FileOutputStream(newDocxFile).use { fos ->
                              document.write(fos)
                          }
                          document.close()
                      }
                      Log.d("MainActivity", "File generated: ${newDocxFile.absolutePath}")

                  }
                  withContext(Dispatchers.Main) {
                      Toast.makeText(
                          this@MainActivity,
                          "Files generated successfully",
                          Toast.LENGTH_SHORT
                      ).show()
                  }


              } catch (e: IOException) {
                  e.printStackTrace()
                  Toast.makeText(this@MainActivity, "Error: ${e.message}", Toast.LENGTH_SHORT).show()
              }
          }
*/




