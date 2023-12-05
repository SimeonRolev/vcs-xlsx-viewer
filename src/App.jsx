import React, { useState } from 'react'
import './App.css'
import Viewer from './Viewer'

function App() {
  const [arrayBuffer, setArrayBuffer] = useState(null)

  React.useEffect(() => {
    fetch('public/Door Schedule w_ Images.xlsx')
      .then(result => result.blob())
      .then(blob => setArrayBuffer(blob.arrayBuffer()))
  }, [])

  return (
    <Viewer arrayBuffer={arrayBuffer} />
  )
}

export default App
