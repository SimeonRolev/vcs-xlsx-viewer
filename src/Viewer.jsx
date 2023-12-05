import React, { useRef } from 'react'
import PropTypes from 'prop-types'
import styled from 'styled-components';
import { renderXlsx } from './original';

const S = {
  Wrapper: styled.div`
      /** xlsx viewer style **/
      .xlsx-viewer-container {
        .xlsx-viewer-tip {
          display: flex;
          justify-content: center;
          align-items: center;
          min-height: 50px;
        }
        .xlsx-viewer-table {
          .xlsx-viewer-table-content {
            width: 100%;
            overflow-x: auto;
            table {
              border-collapse: collapse;
              border-spacing: 0;
              thead {
                background-color: #f0f0f0;
                tr th {
                  border: 1px solid #ccc;
                  padding: 5px;
                  text-align: center;
                }
              }
              tbody {
                background-color: #fff;
                tr td {
                  border: 1px solid #ccc;
                  padding: 5px;
                  text-align: left;
                  vertical-align: middle;
                  &:first-of-type {
                    text-align: center;
                  }
                }
              }
            }
          }
        }
      }

  `
}

function Viewer({ arrayBuffer }) {
  const ref = useRef();

  React.useEffect(() => {
    arrayBuffer && renderXlsx({
      arrayBuffer,
      node: ref.current,
      options: {
        initialSheetIndex: 0
      }
    })
  }, [arrayBuffer])

  return (
    <S.Wrapper ref={ref}>

    </S.Wrapper>
  )
}

Viewer.propTypes = {
  arrayBuffer: PropTypes.object
}

export default Viewer
