// Initial Data
// Load state from LocalStorage
const loadState = () => {
  try {
    const serializedState = localStorage.getItem('state');
    if (serializedState === null) {
      return {};
    }
    return JSON.parse(serializedState);
  } catch (err) {
    return {};
  }
};
// Load styles from LocalStorage
const loadStyles = () => {
  try {
    const serializedState = localStorage.getItem('style');
    if (serializedState === null) {
      return {};
    }
    return JSON.parse(serializedState);
  } catch (err) {
    return {};
  }
};

//STORE
//DataValue
let state = loadState();
//DataKey
let keys = Object.keys(state);
//SelectedValue
let tempo = '';
//DataStyle
let styles = loadStyles();

// UI Components
const refreshButton = `<div style="position: fixed;bottom: 12px;right: 12px;">
                            <Button style="background-color : #31B0D5;color: white;padding: 10px 20px;border-radius: 4px;border-color: #46b8da;"
                                    onclick="refresh();"
                            >
                               <b>Refresh</b> 
                            </Button>
                         </div>`;
const customModal = (content) =>
  `<div id="refreshModal" style="background: white;padding: 20px;position: fixed;z-index: 1; left: 50%;top: 30%;border: 4px solid;">
                        <span onclick="$('#refreshModal').remove()" style="position: absolute;top: 0;right: 9px;cursor:pointer;">&times;</span>
                            ${content}
                      </div>`
const editBar = `<div style="margin-left: 10px; padding: 5px"><b>Edit Bar</b> <span>(Function support : SUM(x:x) AVERAGE(x:x) COUNT(x:x)) MIN(x:x) MAX(x:x)</span>
                    <div>
                        <button class="edit_button" id="bold" disabled><b>Bold</b></button>
                        <button class="edit_button" id="italic" disabled><b>Italic</b></button>
                        <button class="edit_button" id="underline" disabled><b>Underline</b></button>
                    </div>
                 </div>`;


//Update state into localStorage
const localStorageUpdate = (state) => {
  try {

    const serializedState = JSON.stringify(state);
    return (
      localStorage.setItem('state', serializedState)
    );
  } catch (err) {
    return undefined;
  }
}


//Update styles into localStorage
const localStorageStyle = (state) => {
  try {

    const serializedState = JSON.stringify(state);
    return (
      localStorage.setItem('style', serializedState)
    );
  } catch (err) {
    return undefined;
  }
}

// Covert 0-100 to  A-ZZ
const convertToNumberingScheme = (number) => {
  let baseChar = ("A").charCodeAt(0);
  let letters = "";
  do {
    number -= 1;
    letters = String.fromCharCode(baseChar + (number % 26)) + letters;
    number = Math.floor(number / 26);
  } while (number > 0);

  return letters;
}

// Build Grid
const buildGrid = () => {
  for (let i = 0; i < 101; i++) {
    let row = $('#spreadsheet')[0].insertRow();
    for (let j = 0; j < 101; j++) {

      let letter = convertToNumberingScheme(j);
      let cell = row.insertCell();
      cell.setAttribute('contenteditable', true);
      if (i && j) {
        cell.id = letter + i
      }
      else if (i !== 0 || j !== 0) {
        cell.outerHTML = `<th>${i || letter}</th>`;
      }
      else {
        cell.outerHTML = `<th></th>`;
      }
    }
  }
}

// Calculate computed value
const calculateComputed = (key, value) => {
  let result = '';
  try {
    const expression = value.substr(1).toUpperCase();
    let parsed = parse(expression);
    //REGEX Validation
    const functionREGEX = /[a-zA-Z]+\(+[0-9]*\.*[0-9]+\:+([0-9]*\.*[0-9]*\:*[0-9])*\)/g;
    const digitalREGEX = /[0-9]+([.]{1}[0-9]+){0,1}/g;
    const sumREGEX = /sum\(+[0-9]*\.*[0-9]+\:+([0-9]*\.*[0-9]*\:*[0-9])*\)/gi;
    const averageREGEX = /average\(+[0-9]*\.*[0-9]+\:+([0-9]*\.*[0-9]*\:*[0-9])*\)/gi;
    const countREGEX = /count+\(+[0-9]*\.*[0-9]+\:+([0-9]*\.*[0-9]*\:*[0-9])*\)/gi;
    const minREGEX = /min+\(+[0-9]*\.*[0-9]+\:+([0-9]*\.*[0-9]*\:*[0-9])*\)/gi;
    const maxREGEX = /max+\(+[0-9]*\.*[0-9]+\:+([0-9]*\.*[0-9]*\:*[0-9])*\)/gi;

    if (parsed.match(functionREGEX)) {
// FUNCTION SUM(A1:B2)
      if (parsed.match(sumREGEX)) {
        const reducer = (accumulator, currentValue) => parseInt(accumulator) + parseInt(currentValue);

        parsed = parsed.replace(sumREGEX, function (expression) {
          return expression.match(digitalREGEX).reduce(reducer)
        })
      }
// FUNCTION AVERAGE(A1:B2)
      if (parsed.match(averageREGEX)) {
        const reducer = (accumulator, currentValue) => parseInt(accumulator) + parseInt(currentValue);

        parsed = parsed.replace(averageREGEX, function (expression) {
          return expression.match(digitalREGEX).reduce(reducer) / expression.match(digitalREGEX).length
        })

      }
// FUNCTION COUNT(A1:B2)
      if (parsed.match(countREGEX)) {
        parsed = parsed.replace(countREGEX, function (expression) {
          return expression.match(digitalREGEX).length
        })

      }
// FUNCTION MIN(A1:B2)
      if (parsed.match(minREGEX)) {
        parsed = parsed.replace(minREGEX, function (expression) {
          return Math.min(...expression.match(digitalREGEX))
        })
      }

// FUNCTION MAX(A1:B2)
      if (parsed.match(maxREGEX)) {
        parsed = parsed.replace(maxREGEX, function (expression) {
          return Math.max(...expression.match(digitalREGEX))
        })
      }

    }
    result = eval(parsed);

  }
  catch (err) {
    console.error(err)
    $('#spreadsheet').append(customModal('Invalid data entry'));
  }
  if (result || result === 0) {
    state[key]['computed'] = result;

    return result
  }
}

// ReCalculate the computed value of state
const reCalculateComputed = () => {
  keys = Object.keys(state)
  keys.map(key => 'computed' in state[key] ? state[key].computed = calculateComputed(key, state[key].value) : null)
}

// RenderData
const render = () => {
  keys = Object.keys(state)
  keys.map(key => ($(`#${key}`).text('computed' in state[key] ? state[key].computed : state[key].value)))
  //Render format
  const styleKeys = Object.keys(styles)
  if (!!styleKeys) {
    for (let key of styleKeys) {
      const styleNames = Object.keys(styles[key]);
      const styleValues = Object.values(styles[key]);
      for (let i = 0; i < styleNames.length; i++) {
        $(`#${key}`).css(styleNames[i], styleValues[i])
      }
    }
  }
}

// Refresh
const refresh = () => {
  $('#spreadsheet').empty()
  // Initial Data State
  state = loadState();
  // Render Table
  buildGrid();
  if (!!keys) {
    render()
    renderStyles()
    $('#spreadsheet').append(refreshButton);
    $('#spreadsheet').append(customModal('Refreshed'));

    watcher()
  }
}

// Parse formula key to value

const REGEX = /[A-Z]\w*[0-9]/g;

const parse = (expression) => expression.replace(REGEX, function (key) {
  try {
    return 'computed' in state[key] ? state[key].computed : state[key].value
  }
  catch (err) {
    $('#spreadsheet').append(customModal('Invalid data entry'));

  }
})

//Bold controller
const boldController = () => {
  $('#bold').attr('onclick', `
          console.log("!!!")

        if ($('#${tempo}').css('font-weight') == 400) {
        
            $('#${tempo}').css('font-weight', 'bold');
            if(! styles['${tempo}']){
              styles['${tempo}'] = {}
            }
            styles['${tempo}']['font-weight'] = 'bold'
            localStorageStyle(styles)

          }
        else {
        
            $('#${tempo}').css('font-weight', '400');
            if(! styles['${tempo}']){
              styles['${tempo}'] = {}
            }
            styles['${tempo}']['font-weight'] = '400'
            localStorageStyle(styles)

          }
        `)
}
//Italic controller
const italicController = () => {
  $('#italic').attr('onclick', `
        if ($('#${tempo}').css('font-style') == 'normal') {
        
            $('#${tempo}').css('font-style', 'italic');
            if(! styles['${tempo}']){
              styles['${tempo}'] = {}
            }
            styles['${tempo}']['font-style'] = 'italic'
            localStorageStyle(styles)

          }
        else {
            $('#${tempo}').css('font-style', 'normal');
            if(! styles['${tempo}']){
              styles['${tempo}'] = {}
            }
            styles['${tempo}']['font-style'] = 'normal'
            localStorageStyle(styles)          
            }
        `)
}
// Underline Controller
const underlineController = () =>
  $('#underline').attr('onclick', `
        if ($('#${tempo}').css('text-decoration').includes('none')) {
        
            $('#${tempo}').css('text-decoration', 'underline');
            if(! styles['${tempo}']){
              styles['${tempo}'] = {}
            }
            styles['${tempo}']['text-decoration'] = 'underline'
            localStorageStyle(styles)

          }
        else {
        
            $('#${tempo}').css('text-decoration', 'none');
            $('#${tempo}').css('text-decoration', 'none');
            if(! styles['${tempo}']){
              styles['${tempo}'] = {}
            }
            styles['${tempo}']['text-decoration'] = 'none'
            localStorageStyle(styles)

          }
        `)


// Render static grid stySheet
const renderStyles = () => {
  $('#spreadsheet').css({'widith': '100%', 'border-collapse': 'collapse'});
  $('#spreadsheet td, #spreadsheet th').css({
    'border': '1px solid #ddd',
    'padding': '8px',
    'min-width': '60px',
    'max-width': '60px',
    'max-height': '10px',
    'outline': 'none',
  });
  $('#spreadsheet th').css({
    'padding-top': '12px',
    'padding-bottom': '12px',
    'text-align': 'center',
    'background-color': '#4CAF50',
    'color': 'white',
  });
};


//Watch User Interaction
const watcher = () => {
  $("#spreadsheet td")
    .on("blur", function () {
      if ($(this).text() || $(this).text() === 0) {
        $(this).text($(this).text().trim())
        state[$(this).attr('id')] = {}
        state[$(this).attr('id')]['value'] = $(this).text();
        if ($(this).text().startsWith('=')) {
          let computed = calculateComputed($(this).attr('id'), $(this).text())
          $(this).text(computed)
        }
        reCalculateComputed()
        render()
        localStorageUpdate(state)
      }
      else {
        delete state[$(this).attr('id')]
        reCalculateComputed();
        render();
        localStorageUpdate(state)
      }
    })
    .on("focus", function () {
      const preTempo = tempo

      if (!!preTempo) {
        $(`#${preTempo}`).css({
          'border': '1px solid rgb(221, 221, 221)',
          'box-shadow': 'none',
        })
      }
      if ($(this).text() || $(this).text() === 0) {
        $(this).text(state[$(this).attr('id')]['value'])
        tempo = $(this).attr('id')

        $(`#${tempo}`).css({
          'border': '3px solid rgb(224, 135, 135)',
          'box-shadow': 'rgba(255, 0, 0, 0.5) 0px 0px 15px',
        })
        $('.edit_button').removeAttr("disabled")

        // Bold Controller
        boldController()
        //  Italic Controller
        italicController()
        // Underline Controller
        underlineController()

      }
      else {
        tempo = ''
        $('.edit_button').attr("disabled", true)

      }
    })
    .on("keyup", function () {
      const preTempo = tempo
      tempo = $(this).attr('id')

      if (!!preTempo) {
        $(`#${preTempo}`).css({
          'border': '1px solid rgb(221, 221, 221)',
          'box-shadow': 'none',
        })
      }
      $(`#${tempo}`).css({
        'border': '3px solid rgb(224, 135, 135)',
        'box-shadow': 'rgba(255, 0, 0, 0.5) 0px 0px 15px',
      })
      if ($(this).text() || $(this).text() === 0) {

        $('.edit_button').removeAttr("disabled")
      }

      if ($(this).text() || $(this).text() === 0) {
        state[$(this).attr('id')] = {}
        state[$(this).attr('id')]['value'] = $(this).text();

        localStorageUpdate(state)
      }
    })
}


$(document).ready(function () {

  // Render Table
  buildGrid();
  if (!!keys) {
    // Render Data
    render()
  }
  // Add Style
  renderStyles();
  // Render RefreshButton
  $('#spreadsheet').append(refreshButton);

  // Render EditBar
  $('#spreadsheet').before(editBar);

  // Start Watch
  watcher();

})