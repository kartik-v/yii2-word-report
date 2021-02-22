<?php

/**
 * @package   yii2-word-report
 * @author    Kartik Visweswaran <kartikv2@gmail.com>
 * @copyright Copyright &copy; Kartik Visweswaran, Krajee.com, 2021
 * @version   1.0.0
 */

namespace kartik\wordreport\utils;

use yii\base\Component;

/**
 * This Command library helps managing and executing shell commands. Modified version of the
 * Command class from mikehaertl/php-shellcommand.
 */
class Command extends Component
{
    /**
     * @var bool whether to escape any argument passed through `addArg()`.
     * Default is `true`.
     */
    public $escapeArgs = true;

    /**
     * @var bool whether to escape the command passed to `setCommand()` or the
     * constructor.  This is only useful if `$escapeArgs` is `false`. Default
     * is `false`.
     */
    public $escapeCommand = false;

    /**
     * @var bool whether to use `exec()` instead of `proc_open()`. This can be
     * used on Windows system to workaround some quirks there. Note, that any
     * errors from your command will be output directly to the PHP output
     * stream. `getStdErr()` will also not work anymore and thus you also won't
     * get the error output from `getError()` in this case. You also can't pass
     * any environment variables to the command if this is enabled. Default is
     * `false`.
     */
    public $useExec = false;

    /**
     * @var bool whether to capture stderr (2>&1) when `useExec` is true. This
     * will try to redirect the stderr to stdout and provide the complete
     * output of both in `getStdErr()` and `getError()`.  Default is `true`.
     */
    public $captureStdErr = true;

    /**
     * @var string|null the initial working dir for `proc_open()`. Default is
     * `null` for current PHP working dir.
     */
    public $procCwd;

    /**
     * @var array|null an array with environment variables to pass to
     * `proc_open()`. Default is `null` for none.
     */
    public $procEnv;

    /**
     * @var array|null an array of other_options for `proc_open()`. Default is
     * `null` for none.
     */
    public $procOptions;

    /**
     * @var bool|null whether to set the stdin/stdout/stderr streams to
     * non-blocking mode when `proc_open()` is used. This allows to have huge
     * inputs/outputs without making the process hang. The default is `null`
     * which will enable the feature on Non-Windows systems. Set it to `true`
     * or `false` to manually enable/disable it. It does not work on Windows.
     */
    public $nonBlockingMode;

    /**
     * @var int the time in seconds after which a command should be terminated.
     * This only works in non-blocking mode. Default is `null` which means the
     * process is never terminated.
     */
    public $timeout;

    /**
     * @var string pre execution scripts that will be executed before the command
     * and separated by && (only for non windows)
     */
    public $preExecuteScript;

    /**
     * @var string post execution scripts that will be executed after the command
     * and separated by && (only for non windows)
     */
    public $postExecuteScript;

    /**
     * @var null|string the locale to temporarily set before calling
     * `escapeshellargs()`. Default is `null` for none.
     */
    public $locale;

    /**
     * @var null|string|resource to pipe to standard input
     */
    protected $_stdIn;

    /**
     * @var string the command to execute
     */
    protected $_command;

    /**
     * @var array the list of command arguments
     */
    protected $_args = [];

    /**
     * @var string the full command string to execute
     */
    protected $_execCommand;

    /**
     * @var string the stdout output
     */
    protected $_stdOut = '';

    /**
     * @var string the stderr output
     */
    protected $_stdErr = '';

    /**
     * @var int the exit code
     */
    protected $_exitCode;

    /**
     * @var string the error message
     */
    protected $_error = '';

    /**
     * @var bool whether the command was successfully executed
     */
    protected $_executed = false;

    /**
     * @param string|resource $stdIn If set, the string will be piped to the
     * command via standard input. This enables the same functionality as
     * piping on the command line. It can also be a resource like a file
     * handle or a stream in which case its content will be piped into the
     * command like an input redirection.
     * @return static for method chaining
     */
    public function setStdIn($stdIn)
    {
        $this->_stdIn = $stdIn;
        return $this;
    }

    /**
     * @return string|null the command that was set through setCommand() or
     * passed to the constructor. `null` if none.
     */
    public function getCommand()
    {
        return $this->_command;
    }

    /**
     * @param string $command the command or full command string to execute,
     * like 'gzip' or 'gzip -d'.  You can still call addArg() to add more
     * arguments to the command. If $escapeCommand was set to true, the command
     * gets escaped with escapeshellcmd().
     * @return static for method chaining
     */
    public function setCommand($command)
    {
        if ($this->escapeCommand) {
            $command = escapeshellcmd($command);
        }
        if ($this->getIsWindows()) {
            // Make sure to switch to correct drive like "E:" first if we have
            // a full path in command
            if (isset($command[1]) && $command[1] === ':') {
                $position = 1;
                // Could be a quoted absolute path because of spaces.
                // i.e. "C:\Program Files (x86)\file.exe"
            } elseif (isset($command[2]) && $command[2] === ':') {
                $position = 2;
            } else {
                $position = false;
            }

            // Absolute path. If it's a relative path, let it slide.
            if ($position) {
                $command = sprintf(
                    $command[$position - 1] . ': && cd %s && %s',
                    escapeshellarg(dirname($command)),
                    escapeshellarg(basename($command))
                );
            }
        }
        $this->_command = $command;
        return $this;
    }

    /**
     * @return string|bool the full command string to execute. If no command
     * was set with setCommand() or passed to the constructor it will return
     * `false`.
     */
    public function getExecCommand()
    {
        if ($this->_execCommand === null) {
            $command = $this->getCommand();
            if (!$command) {
                $this->_error = 'Could not locate any executable command';
                return false;
            }
            if (!is_executable($command)) {
                $this->_error = "The '{$command}' command was not found on server or is not executable by the current user!";
                return false;
            }
            $args = $this->getArgs();
            $this->_execCommand = $args ? $command . ' ' . $args : $command;
        }
        if (!empty($this->preExecuteScript)) {
            $this->_execCommand = $this->preExecuteScript . ' && ' . $this->_execCommand;
        }
        if (!empty($this->postExecuteScript)) {
            $this->_execCommand .= ' && ' . $this->postExecuteScript;
        }
        return $this->_execCommand;
    }

    /**
     * @return string the command args that where set with setArgs() or added
     * with addArg() separated by spaces
     */
    public function getArgs()
    {
        return implode(' ', $this->_args);
    }

    /**
     * @param string $args the command arguments as string. Note that these
     * will not get escaped!
     * @return static for method chaining
     */
    public function setArgs($args)
    {
        $this->_args = [$args];
        return $this;
    }

    /**
     * @param string $key the argument key to add e.g. `--feature` or
     * `--name=`. If the key does not end with and `=`, the $value will be
     * separated by a space, if any. Keys are not escaped unless $value is null
     * and $escape is `true`.
     * @param string|array|null $value the optional argument value which will
     * get escaped if $escapeArgs is true.  An array can be passed to add more
     * than one value for a key, e.g. `addArg('--exclude',
     * array('val1','val2'))` which will create the option `'--exclude' 'val1'
     * 'val2'`.
     * @param bool|null $escape if set, this overrides the $escapeArgs setting
     * and enforces escaping/no escaping
     * @return static for method chaining
     */
    public function addArg($key, $value = null, $escape = null)
    {
        $doEscape = $escape !== null ? $escape : $this->escapeArgs;
        $useLocale = $doEscape && $this->locale !== null;

        if ($useLocale) {
            $locale = setlocale(LC_CTYPE, 0);   // Returns current locale setting
            setlocale(LC_CTYPE, $this->locale);
        }
        if ($value === null) {
            $this->_args[] = $doEscape ? escapeshellarg($key) : $key;
        } else {
            if (substr($key, -1) === '=') {
                $separator = '=';
                $argKey = substr($key, 0, -1);
            } else {
                $separator = ' ';
                $argKey = $key;
            }
            $argKey = $doEscape ? escapeshellarg($argKey) : $argKey;

            if (is_array($value)) {
                $params = [];
                foreach ($value as $v) {
                    $params[] = $doEscape ? escapeshellarg($v) : $v;
                }
                $this->_args[] = $argKey . $separator . implode(' ', $params);
            } else {
                $this->_args[] = $argKey . $separator .
                    ($doEscape ? escapeshellarg($value) : $value);
            }
        }
        if ($useLocale) {
            setlocale(LC_CTYPE, $locale);
        }

        return $this;
    }

    /**
     * @param bool $trim whether to `trim()` the return value. The default is `true`.
     * @return string the command output (stdout). Empty if none.
     */
    public function getOutput($trim = true)
    {
        return $trim ? trim($this->_stdOut) : $this->_stdOut;
    }

    /**
     * @param bool $trim whether to `trim()` the return value. The default is `true`.
     * @return string the error message, either stderr or an internal message.
     * Empty string if none.
     */
    public function getError($trim = true)
    {
        return $trim ? trim($this->_error) : $this->_error;
    }

    /**
     * @param bool $trim whether to `trim()` the return value. The default is `true`.
     * @return string the stderr output. Empty if none.
     */
    public function getStdErr($trim = true)
    {
        return $trim ? trim($this->_stdErr) : $this->_stdErr;
    }

    /**
     * @return int|null the exit code or null if command was not executed yet
     */
    public function getExitCode()
    {
        return $this->_exitCode;
    }

    /**
     * @return bool whether the command was successfully executed
     */
    public function getIsExecuted()
    {
        return $this->_executed;
    }

    /**
     * Execute the command
     *
     * @return bool whether execution was successful. If `false`, error details
     * can be obtained from getError(), getStdErr() and getExitCode().
     */
    public function execute()
    {
        $command = $this->getExecCommand();

        if (!$command) {
            return false;
        }
        if ($this->useExec) {
            $execCommand = $this->captureStdErr ? "$command 2>&1" : $command;
            exec($execCommand, $output, $this->_exitCode);
            $this->_stdOut = implode("\n", $output);
            if ($this->_exitCode !== 0) {
                $this->_stdErr = $this->_stdOut;
                $this->_error = empty($this->_stdErr) ? 'Command failed' : $this->_stdErr;
                return false;
            }
        } else {
            $isInputStream = $this->_stdIn !== null &&
                is_resource($this->_stdIn) &&
                in_array(get_resource_type($this->_stdIn), ['file', 'stream']);
            $isInputString = is_string($this->_stdIn);
            $hasInput = $isInputStream || $isInputString;
            $hasTimeout = $this->timeout !== null && $this->timeout > 0;

            $descriptors = [
                1 => ['pipe', 'w'],
                2 => ['pipe', $this->getIsWindows() ? 'a' : 'w'],
            ];
            if ($hasInput) {
                $descriptors[0] = ['pipe', 'r'];
            }


            // Issue #20 Set non-blocking mode to fix hanging processes
            $nonBlocking = $this->nonBlockingMode === null ?
                !$this->getIsWindows() : $this->nonBlockingMode;

            $startTime = $hasTimeout ? time() : 0;
            $process = proc_open($command, $descriptors, $pipes, $this->procCwd, $this->procEnv, $this->procOptions);

            if (is_resource($process)) {
                if ($nonBlocking) {
                    stream_set_blocking($pipes[1], false);
                    stream_set_blocking($pipes[2], false);
                    if ($hasInput) {
                        $writtenBytes = 0;
                        $isInputOpen = true;
                        stream_set_blocking($pipes[0], false);
                        if ($isInputStream) {
                            stream_set_blocking($this->_stdIn, false);
                        }
                    }

                    // Due to the non-blocking streams we now have to check in
                    // a loop if the process is still running. We also need to
                    // ensure that all the pipes are written/read alternately
                    // until there's nothing left to write/read.
                    $isRunning = true;
                    while ($isRunning) {
                        $status = proc_get_status($process);
                        $isRunning = $status['running'];

                        // We first write to stdIn if we have an input. For big
                        // inputs it will only write until the input buffer of
                        // the command is full (the command may now wait that
                        // we read the output buffers - see below). So we may
                        // have to continue writing in another cycle.
                        //
                        // After everything is written it's safe to close the
                        // input pipe.
                        if ($isRunning && $hasInput && $isInputOpen) {
                            if ($isInputStream) {
                                $written = stream_copy_to_stream($this->_stdIn, $pipes[0], 16 * 1024, $writtenBytes);
                                if ($written === false || $written === 0) {
                                    $isInputOpen = false;
                                    fclose($pipes[0]);
                                } else {
                                    $writtenBytes += $written;
                                }
                            } else {
                                if ($writtenBytes < strlen($this->_stdIn)) {
                                    $writtenBytes += fwrite($pipes[0], substr($this->_stdIn, $writtenBytes));
                                } else {
                                    $isInputOpen = false;
                                    fclose($pipes[0]);
                                }
                            }
                        }

                        // Read out the output buffers because if they are full
                        // the command may block execution. We do this even if
                        // $isRunning is `false`, because there could be output
                        // left in the buffers.
                        //
                        // The latter is only an assumption and needs to be
                        // verified - but it does not hurt either and works as
                        // expected.
                        //
                        while (($out = fgets($pipes[1])) !== false) {
                            $this->_stdOut .= $out;
                        }
                        while (($err = fgets($pipes[2])) !== false) {
                            $this->_stdErr .= $err;
                        }

                        $runTime = $hasTimeout ? time() - $startTime : 0;
                        if ($isRunning && $hasTimeout && $runTime >= $this->timeout) {
                            // Only send a SIGTERM and handle status in the next cycle
                            proc_terminate($process);
                        }

                        if (!$isRunning) {
                            $this->_exitCode = $status['exitcode'];
                            if ($this->_exitCode !== 0 && empty($this->_stdErr)) {
                                if ($status['stopped']) {
                                    $signal = $status['stopsig'];
                                    $this->_stdErr = "Command stopped by signal $signal";
                                } elseif ($status['signaled']) {
                                    $signal = $status['termsig'];
                                    $this->_stdErr = "Command terminated by signal $signal";
                                } else {
                                    $this->_stdErr = 'Command unexpectedly terminated without error message';
                                }
                            }
                            fclose($pipes[1]);
                            fclose($pipes[2]);
                            proc_close($process);
                        } else {
                            // The command is still running. Let's wait some
                            // time before we start the next cycle.
                            usleep(10000);
                        }
                    }
                } else {
                    if ($hasInput) {
                        if ($isInputStream) {
                            stream_copy_to_stream($this->_stdIn, $pipes[0]);
                        } elseif ($isInputString) {
                            fwrite($pipes[0], $this->_stdIn);
                        }
                        fclose($pipes[0]);
                    }
                    $this->_stdOut = stream_get_contents($pipes[1]);
                    $this->_stdErr = stream_get_contents($pipes[2]);
                    fclose($pipes[1]);
                    fclose($pipes[2]);
                    $this->_exitCode = proc_close($process);
                }

                if ($this->_exitCode !== 0) {
                    $this->_error = $this->_stdErr ?
                        $this->_stdErr :
                        "Failed without error message: $command (Exit code: {$this->_exitCode})";
                    return false;
                }
            } else {
                $this->_error = "Could not run command $command";
                return false;
            }
        }

        $this->_executed = true;

        return true;
    }

    /**
     * @return bool whether we are on a Windows OS
     */
    public function getIsWindows()
    {
        return strncasecmp(PHP_OS, 'WIN', 3) === 0;
    }
}